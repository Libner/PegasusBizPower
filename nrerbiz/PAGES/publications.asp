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

Sub PageTitleVis(id ,isvis , sort )
	If Trim(sort) = "page" Then
	    sql = "update pagesTav set Page_Visible=" & isvis & " WHERE Page_Id=" & id
	ElseIf Trim(sort) = "title" Then
	    sql = "update pagesTav set Page_Visible_Title=" & isvis & " WHERE Page_Id=" & id
	End If
	con.execute sql
End Sub

Sub UpAndDownPages(dir, ctd, place, p)
    Dim stam
    Dim pl
    pl = 0
    
    If Not IsNull(place) Then
       pl = place
    End If
    
    If dir = "down" Then
        stam = CInt(pl) + 1
    ElseIf dir = "up" Then
        stam = CInt(pl) - 1
    End If
    
    sql = "update pagesTav set Page_Order=-10 WHERE Page_Order=" & CStr(stam) & " and Category_Id=" & CStr(ctd) & " and Page_Parent=" & CStr(p)
    con.execute sql
    sql1 = "update pagesTav set Page_Order=" & (CStr(stam) & " WHERE Page_Order=" & CStr(pl)) & " and Category_Id=" & CStr(ctd) & " and Page_Parent=" & CStr(p)
    con.execute sql1
    sql2 = "update pagesTav set Page_Order=" & CStr(pl) & " WHERE Page_Order=-10 "
    con.execute sql2
End Sub

Sub OrderPage(pid,cid,cn)
Dim p 
    sql = "select Page_Order,Page_Id,Category_Id from pagesTav where Category_Id=" & cid & " and Page_Parent=0 order by Page_Order "
    Set myNum = con.execute(sql)
    If Not myNum.EOF Then
		 p = 1
		 Dim pida
		 Do While p <= cn
			pida = myNum("Page_Id")
		    sql = "update pagesTav set Page_Order=" & p & " where Page_Id=" & pida
		    con.execute sql
		    myNum.MoveNext
		    p = p + 1
		Loop
   End If
   myNum.Close
   Set myNum = Nothing
End Sub

%>
<html>

<head>
<title>ניהול תבניות תוכן</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
<link href="../dynamic_style.css" rel="STYLESHEET" type="text/css">
</head>

<script LANGUAGE="JavaScript">
<!--
function CheckDel() {
  return (confirm("?האם ברצונך למחוק את הדף"))    
}
<!--End-->
function CheckProcess() {
  return (confirm("?האם ברצונך להעביר לארכיון"))    
}
function  updvisibility(pageid,isvis)
{
	document.ff.pageidvis.value=pageid
	document.ff.isvisible.value=isvis
	document.ff.submit()
}

function  updvisibility_title(pageid,isvis)
{
	document.ff.titlevis.value=pageid
	document.ff.isvisible_title.value=isvis
	document.ff.submit()
}

function mapwindow(mainid,catid,pageid)
{
	url="copypage.asp?mainid="+mainid+"&catid=" + catid+"&pageid=" + pageid
	cpagewin=window.open(url,"cpagewin","left=100,top=50,height=400,width=500,status=no,scrollbars=yes,toolbar=no,menubar=no,location=no,resizable=yes");
	return false
}

function mapwindowCopy(mainid,catid,pageid)
{
	url="copypageDubl.asp?mainid="+mainid+"&catid=" + catid+"&pageid=" + pageid
	cpagewin=window.open(url,"cpagewin","left=100,top=50,height=400,width=500,status=no,scrollbars=yes,toolbar=no,menubar=no,location=no,resizable=yes");
	return false
}

<!--End-->
</script>  

<%
sqlstring="SELECT * from workers where loginName='" & userName & "' and password='"& password &"' "
set worker=con.execute(sqlstring)
if worker.EOF then%>
<p><center><font color="red" size="4"><b><%=permissions%></b></font></center>
<%else
  session("admin.username")     =  username
  session("admin.password")     =  password
  siteid=session("siteid")
  set fs                  =  server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
  isRightMenu             =  Request.querystring("isright")
  place                   =  Request.QueryString("place")
  categId                 =  Request.QueryString("catId")
  pageId                  =  request.querystring("pageID")
  parent                  =  request.querystring("parentID")
  innerparent             =  request("innerparent")
  subcat                  =  request("subcat")
  home                    =  Request.QueryString("home")
  linkhome                =  Request.QueryString("linkhome")
  isvisible               =  request("isvisible")
  isvisible_title         =  request("isvisible_title")

'INITIALIZATION OF THE OBJECT.................................................

if isvisible<>"" then 
  PageTitleVis request("pageidvis"),isvisible,"page"
end if
if isvisible_title<>"" then 
  PageTitleVis request("titlevis"),isvisible_title,"title"
end if

'May be not used now-------------------------------------------------------------
if home<>"" and  home<>nil then
   con.execute ("update pagesTav set Page_Home=0 WHERE Page_Home=1 and Category_Id="&categId & " and Page_Parent=" & parent)
   con.execute ("update pagesTav set Page_Home=1 WHERE Page_ID=" & pageId &" and Category_Id="&categId & " and Page_Parent=" & parent)
end if 
if linkhome<>"" and  linkhome<>nil then
   con.execute ("update pagesTav set Page_Link_Home=1-Page_Link_Home WHERE Page_ID=" & pageId &" and Category_Id="&categId)
end if 
'-------------------------------------------------------------------------------------
 
'Delete Page
if Request.QueryString("DEL")="1" then 
    '//files 
    '//links
	'//elements
	'//order
	sql1 = "SELECT Page_Order,Page_Home,Page_Parent from pagesTav where Page_Id=" & pageId
    Set dord = con.execute(sql1)
		if not dord.eof then
			parent = dord("Page_Parent")
			delord = dord("Page_Order")
			delhome = dord("Page_Home")
		end if	
    Set dord = Nothing
   'del
   recs = 0
   sql2 = "DELETE FROM pagesTav WHERE Page_Id=" & pageId
   con.execute sql2, recs
    'upd after del
    if recs<>0 then
		sql3 = "update pagesTav set Page_Order=Page_Order-1 WHERE Page_Order>" & delord & " and Category_Id=" & categId & " and Page_Parent=" & parent
		con.execute sql3
		If CBool(delhome) = True Then
		   sql4 = "update pagesTav set Page_Home=1 WHERE Page_Order=1 and Category_Id=" & categId & " and Page_Parent=" & parent
		   con.execute sql4
		End If
    end if
    sql5 = "Delete from pagesTav where Page_Parent=" & pageId
    con.execute sql5
end if

if Request.QueryString("down")<>nil then
    UpAndDownPages "down",categId,place,parent
Elseif Request.QueryString("up")<>nil then
    UpAndDownPages "up",categId,place,parent
end if
   

    
sql6 = "SELECT * FROM Publication_Categories WHERE Publication_Category_ID=" & categId
Set publ = con.execute(sql6)
catName = publ("Publication_Category_Name")
maincat = publ("Main_Category_ID")
Set publ = Nothing
    
cnt = 0
sql7 = "select count(*) as myCont from pagesTav where Category_Id=" & categId & " and Page_Parent=0"
Set mcnt = con.execute(sql7)
if not mcnt.eof then
   cnt = mcnt("myCont")
end if
Set mcnt = Nothing    
    
OrderPage pageId,categId,cnt
 
sql9 = "SELECT p1.*,isExistInnnerPages = case when Exists (SELECT page_id FROM pagesTav as p2 where page_parent=p1.page_id) then 1 else 0 End FROM pagesTav as p1 Where Category_ID = " & categId & " And Page_Parent = 0 order by Page_Order"
set pr   = con.execute(sql9)
sql10 = "SELECT site_Id,Main_Category_Name FROM Main_Categories WHERE Main_Category_ID=" & maincat
set main = con.execute(sql10)
    
%>
<body class="body_admin">
<div align="right">
<form name="ff" action="publications.asp?subcat=<%=subcat%>&innerparent=<%=page_id%>&catId=<%=categId%>" method="post">
   <input type="hidden" name=isvisible value="">
   <input type="hidden" name="pageidvis" value="">
   <input type="hidden" name=isvisible_title value="">
   <input type="hidden" name="titlevis" value="">
</form>
<table class="table_admin_2" border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
  <%
  sql11 = "SELECT Site_Name FROM Sites WHERE Site_Id=" & SiteId
  set sites = con.execute(sql11)%>
 <tr>
    <td class="a_title_big" width="100%" dir=rtl><center> מנגנון ניהול אתר "<%=sites("Site_Name")%>"</center>
    </td>
  </tr>
<tr>
  <td class="td_admin_4" height="10">
  </td>
</tr>
<tr><td class="title_admin_1" align="right" valign="middle" width="100%" nowrap dir="rtl">&nbsp;<%=main("Main_Category_Name")%></td></tr>
<tr >
   <td class="td_admin_4" align="left"  valign="middle" nowrap>
			<table class="table_admin_2" width="100%" border="0"  cellpadding="1" cellspacing="0">
					<tr>
							<td align="left" valign="middle" nowrap>
								<a  class="button_admin_1" href="../choose.asp">&nbsp;חזרה לניהול קטגוריות ראשיות&nbsp;</a></td>
							<td align="left" valign="middle" nowrap dir="rtl">
								<a class="button_admin_1" href="admin.asp?maincat=<%=maincat%>">&nbsp;חזרה לדף ניהול קטגוריה <%=main("Main_Category_Name")%>&nbsp;</a>
							</td>
							<td width="100%">&nbsp;</td>
					</tr>
			</table>
	</td>
</tr>
<tr><td class="td_admin_4" align="center" valign="middle" nowrap>&nbsp;</td></tr>
</table>
<table class="table_admin_2" width="100%" border="0" cellpadding="0" cellspacing="0">
   <tr><td class="title_admin_1" colspan="7" height="1"></td></tr>
   <tr>
       <td class="title_admin_1" colspan="3" align="left" valign="middle"  nowrap>
	         <a class="button_admin_1" href="AddPage.asp?catId=<%=categId%>&isright=<%=isRightMenu%>&innerparent=<%=page_id%>&subcat=<%=subcat%>&parentID=0">הוספת דף חדש</a></td>
	   <td class="title_admin_1"  colspan="4" align="right"  nowrap dir="rtl">&nbsp;<%=catName%></td>
   </tr>
   <tr><td class="title_admin_1" colspan="7" height="1"></td></tr>
<tr><td  colspan="7"  class="td_line_between"></td></tr>
<%i=0	
do while not pr.EOF
	    pageOrder=pr("Page_Order")
		perName=pr("Page_Title")
	    page_id=pr("Page_ID")
	    parent_id=pr("Page_Parent")
	    pageURL=pr("Page_URL")
	    pagevisible=pr("Page_Visible")
	    titlevis=pr("Page_Visible_Title")
	    pageisHome=pr("Page_Home")
	    pageisLinkHome=pr("Page_Link_Home")
	    perDate=pr("Page_Date")
	    isExistInnnerPages=pr("isExistInnnerPages")
	    i=i+1
	    %>
<tr>
    <td class="td_admin_4_left_align" align="left" valign="top">
	   <table border="0" cellpadding="1" cellspacing="0">
	      <%if not isNull(pageURL) and trim(PageURL)<>"" then%>
	          <tr><td class="td_admin_4" align="center" valign="bottom"><a class="button_edit_1"  href="AddPage.asp?pageId=<%=page_id%>&parentID=0&innerparent=<%=page_id%>&subcat=<%=subcat%>&catId=<%=categId%>">עדכון</a></td></tr>	
	      <%else%>
	      <!--OLD component
	          <tr><td class="td_admin_4" align="center" valign="bottom"><a class="button_edit_1" href="editPageContent.asp?pageId=<%=page_id%>&isright=<%=isRightMenu%>&parentID=0&innerparent=<%=page_id%>&subcat=<%=subcat%>&catId=<%=categId%>&maincat=<%=maincat%>">עדכון</a></td></tr>
	      -->
	          <tr><td class="td_admin_4" align="center" valign="bottom"><a class="button_edit_1" href="editpage.aspx?pageId=<%=page_id%>&isright=<%=isRightMenu%>&parentID=0&innerparent=<%=page_id%>&siteId=<%=siteId%>&catId=<%=categId%>&maincat=<%=maincat%>">עדכון</a></td></tr>
	      <%end if%>     
	      <tr><td class="td_admin_4" align="center" valign="top" ><a class="button_delete_1" href="publications.asp?pageId=<%=page_id%>&isright=<%=isRightMenu%>&parentID=0&innerparent=<%=page_id%>&subcat=<%=subcat%>&catId=<%=categId%>&DEL=1" ONCLICK="return CheckDel()">מחיקה</a></td></tr>
	</table>
	</td>
	<td class="td_admin_4_left_align"  valign="middle">
	<!--
		 <%if pageisHome=true then%>
			<img src="images/Main_lamp_on.gif" border="0" hspace="3">
		 <%else%>
			<a href="publications.asp?pageId=<%=page_id%>&isright=<%=isRightMenu%>&parentID=0&innerparent=<%=page_id%>&subcat=<%=subcat%>&catId=<%=catID%>&home=1">
			    <img src="images/Main_lamp_off.gif" border="0" hspace="3">
			</a>
		 <%end if%>
	-->
	</td>
	<td align="left" valign="middle" class="td_admin_4_left_align">
	   <table border="0" valign="middle" cellpadding="0" cellspacing="0" width="100%">
	      <tr>
		     <td align="left" valign="middle" class="td_admin_4_left_align" nowrap>
		        <table border="0" cellpadding="1" cellspacing="0">
		  <tr>
		     <td colspan="3" class="td_admin_4_left_align"  align="left" valign="middle"nowrap>
			      <%if isNull(pageURL) or trim(PageURL)="" then%> 	
			       	  <a class="button_admin_1" href="AddPage.asp?catId=<%=categId%>&isright=<%=isRightMenu%>&innerparent=<%=page_id%>&subcat=<%=subcat%>&parentID=<%=page_id%>">הוספת דף פנימי</a>
			      <%end if%>
		     </td>
		  </tr>
		           <!--tr>
		              <td align="left" valign="middle" class="td_admin_4_left_align" nowrap>
		                 <a class="button_copy" href="publications.asp?arch=1&pageId=<%=page_id%>&catId=<%=catID%>" onclick="return CheckProcess();">
		                      לארכיון
		                 </a>
		              </td>
       		          <td align="center" valign="middle" class="td_admin_4_left_align" nowrap>
			             <a class="button_copy"  href="" onclick="return mapwindow('<%=maincat%>','<%=catId%>','<%=page_id%>')">
			                 העברה
			             </a>
		              </td>
		              <td align="right" valign="middle" class="td_admin_4_right_align" nowrap>
			              <a  class="button_copy" href="" onclick="return mapwindowCopy('<%=maincat%>','<%=catId%>','<%=page_id%>')">
			                  העתק
			              </a>
		              </td>
		             </tr-->
		        </table>
		        </td>
		              <td align="left" valign="middle" class="td_admin_4_left_align" nowrap>
			                <input type="checkbox" <%if pagevisible<>0 then%> checked onclick="updvisibility(<%=page_id%>,0)" <%else%> onclick="updvisibility(<%=page_id%>,1)" <%end if%>  id=checkbox1 name=checkbox1>&nbsp;הצג תוכן באתר&nbsp;&nbsp;
			                <br><input type="checkbox" <%if titlevis<>0 then%> checked onclick="updvisibility_title(<%=page_id%>,0)" <%else%> onclick="updvisibility_title(<%=page_id%>,1)" <%end if%>  id=checkbox2 name=checkbox2>&nbsp;הצג כותרת באתר&nbsp;&nbsp;
		              </td>
		           </tr>
	</table>
 </td>
 <td align="left" valign="middle" class="td_admin_4_left_align">&nbsp;
 </td>
 <td align="right" valign="middle" class="td_admin_4_right_align" width="100%">
	<table  border="0" cellpadding="0" cellspacing="0"  width="100%">
       <tr>
		  <td  align="right" valign="middle" class="td_admin_4_right_align" width="100%" dir="rtl">
			   <%if isNull(pageURL) or trim(PageURL)="" then%> 	
				<%if isExistInnnerPages=1 then%>
				  <a class="link_categ" href="publications.asp?subcat=<%=page_id%>&innerparent=<%=page_id%>&catId=<%=categId%>" title="ל רשימת דפים פנימיים">
				      <%=perName%>
				  </a>&nbsp;
				 <%else%>
				      <b><%=perName%>&nbsp;</b>
				 <%end if%>
			   <%else%>
				  <a class="link_categ" href="<%=pageURL%>" target="_blank">
				      <%=perName%>
				  </a>
			   <%end if%>
		  </td>
		  <td class="td_admin_4_right_align">
		      <%=perDate%>
		  </td>
		  <td class="td_admin_4_right_align">&nbsp;
		  </td>
	   </tr>
	 </table>
  </td>
  
  <%if i>1 then%>
  <td width="1%" align="center" valign="bottom" class="td_admin_2"><a href="publications.asp?catId=<%=categId%>&isright=<%=isRightMenu%>&innerparent=<%=page_id%>&subcat=<%=subcat%>&parentID=0&up=1&place=<%=i%>"><img src="images/up.gif" alt="להעביר את הטקסט למעלה" border="0" align="bottom" vspace="0"></a></td>
  <%else%>
  <td width="1%" align="center" valign="bottom" class="td_admin_2">
	  &nbsp;
  </td>
  <%end if%>
  <%if i<cnt then%>
  <td width="1%" align="center" valign="top" class="td_admin_2">
	<a href="publications.asp?catId=<%=categId%>&isright=<%=isRightMenu%>&innerparent=<%=page_id%>&subcat=<%=subcat%>&parentID=0&down=1&place=<%=i%>"><img src="images/down.gif" alt="להעביר את הטקסט למטה" border="0" vspace="0"></a>
  </td>
  <%else%>
  <td width="1%" align="center" valign="top" class="td_admin_2">
	  &nbsp;
  </td>
  <%end if%>
</tr>
<tr>
    <td  colspan="7"  class="td_line_between">
    </td>
</tr>
<%if subcat<>nil and subcat<>"" then
if Cint(subcat)=page_id then%>
<!--subcategories-->
<% 
sql12 = "select * from pagesTav WHERE Category_ID = " & categId & " and Page_Parent=" & page_id & "  order by  Page_Order"
set catinner = con.execute(sql12)
if not catinner.EOF then
	sql13 = "SELECT COUNT(*) AS numpage from pagesTav WHERE Category_ID = " & categId & " AND Page_Parent=" & page_id
	Set reca = con.execute(sql13)
		numinside = reca("numpage")
	reca.Close
	Set reca = Nothing
%>
<tr>
    <td class="td_line_between"  colspan="7" height="1">
    </td>
</tr>
<tr>
    <td  class="td_admin_4">
	   &nbsp;
    </td>
    <td  class="td_admin_4_right_align" align="right" colspan="4">
         <table border="0" cellpadding="0" cellspacing="0" width="100%">
<%      
        ii = 0	
        do while not catinner.EOF
        
	    pageOrder       =   catinner("Page_Order")
		perName         =   catinner("Page_Title")
	    page_id         =   catinner("Page_ID")
	    parent_id       =   catinner("Page_Parent")
	    pageURL         =   catinner("Page_URL")
	    pagevisible     =   catinner("Page_Visible")
        titlevis        =   catinner("Page_Visible_Title")
	    pageisHome      =   catinner("Page_Home")
	    pageisLinkHome  =   catinner("Page_Link_Home")
	    perDate         =   catinner("Page_Date")
	    ii              =   ii + 1
%>
<tr>
    
	<td align="left" valign="top"  class="td_admin_4_left_align">
	<table  border="0" cellpadding="0" cellspacing="0">
	<%if not isNull(pageURL)<>nil and trim(PageURL)<>"" then%>
	<tr>
	   <td align="center" valign="bottom" class="td_admin_4">
		  <a class="button_edit_1" href="AddPage.asp?pageId=<%=page_id%>&innerparent=<%=page_id%>&subcat=<%=subcat%>&catId=<%=categId%>">
			    עדכון
		  </a>
	   </td>
	</tr>	
	<%else%>
	<tr>
		<td align="center" valign="bottom" class="td_admin_4">
		   <a class="button_edit_1" href="editPageContent.asp?pageId=<%=page_id%>&innerparent=<%=page_id%>&subcat=<%=subcat%>&catId=<%=categId%>&maincat=<%=maincat%>">
			    עדכון
		   </a>
		</td>
	</tr>
	<%end if%>
	<tr>
	   <td align="center" valign="top" class="td_admin_4">
	       <a class="button_delete_1" href="publications.asp?pageId=<%=page_id%>&innerparent=<%=page_id%>&subcat=<%=subcat%>&catId=<%=categId%>&DEL=1" ONCLICK="return CheckDel()">
	          מחיקה
	       </a>
	   </td>
    </tr>
  </table>
</td>
<td align="left" valign="middle"   nowrap class="td_admin_4_left_align">
   <table  border="0" cellpadding="1" cellspacing="0">
      <tr>
         <!--td class="td_admin_4_left_align" align="left" valign="middle"  nowrap>
            <a class="button_copy" href="publications.asp?arch=1&pageId=<%=page_id%>&catId=<%=catID%>" onclick="return CheckProcess();">
                 לארכיון
            </a>
         </td>
         <td class="td_admin_4_left_align" align="left" valign="middle"  nowrap>
	         <a class="button_copy" href="" onclick="return mapwindow('<%=maincat%>','<%=catId%>','<%=page_id%>')">
	              העברה
	         </a>
       	 </td>
	     <td class="td_admin_4_left_align" align="left" valign="middle" nowrap>
 		     <a class="button_copy" href="" onclick="return mapwindowCopy('<%=maincat%>','<%=catId%>','<%=page_id%>')">
 		          העתק
 		     </a>
	     </td-->
         <td align="left" valign="baseline" class="td_admin_4_left_align" nowrap>
		     <input type="checkbox" <%if pagevisible<>0 then%> checked onclick="updvisibility(<%=page_id%>,0)" <%else%> onclick="updvisibility(<%=page_id%>,1)" <%end if%>  id=checkbox2 name=checkbox2>&nbsp;הצג תוכן באתר
		     <!--br><input type="checkbox" <%if titlevis<>0 then%> checked onclick="updvisibility_title(<%=page_id%>,0)" <%else%> onclick="updvisibility_title(<%=page_id%>,1)" <%end if%>  id=checkbox2 name=checkbox2-->&nbsp;<%'=reverse("הצג כותרת באתר")%>
         </td>
      </tr>
   </table>
</td>
 <td align="left" valign="middle" class="td_admin_4_left_align">&nbsp;
 </td>
<td align="right" valign="middle" class="td_admin_4_right_align"  width="100%">
   <table align="right" border="0" cellpadding="0" cellspacing="0" width="100%" >
	  <tr>
		  <td  align="right" valign="middle"  class="td_admin_4_right_align" width="100%" dir="rtl">
		    <%if not isNull(pageURL)<>nil and trim(PageURL)<>"" then%>
			<a class="link_categ" href="<%=pageURL%>" target="_blank">
			   <%=perName%>
			   &nbsp;
			</a>
			<%else%>
			<p class="page_title">
			<%=perName%>
			&nbsp;</p>
			<%end if%>
		  </td>
		  <td class="td_admin_4_right_align">
		       <%=perDate%>
		  </td>
		  <td class="td_admin_4_right_align">&nbsp;
		  </td>
	 </tr>
   </table>
</td>
	<%if ii>1 then%>
	<td width="1%" align="center" valign="bottom" class="td_admin_2"><a href="publications.asp?catId=<%=categId%>&innerparent=<%=page_id%>&subcat=<%=subcat%>&parentID=<%=parent_id%>&up=1&place=<%=ii%>"><img src="images/up.gif" alt="להעביר את הטקסט למעלה" border="0" ></a></td>
	<%else%>
	<td width="1%" align="center" valign="bottom" class="td_admin_2">
	  &nbsp;
	</td>
	<%end if%>
	<%if ii<numinside then%>
	<td width="1%" align="center" valign="top" class="td_admin_2"><a href="publications.asp?catId=<%=categId%>&innerparent=<%=page_id%>&subcat=<%=subcat%>&parentID=<%=parent_id%>&down=1&place=<%=ii%>"><img src="images/down.gif" alt="להעביר את הטקסט למטה" border="0"></a></td>
	<%else%>
	<td width="1%" align="center" valign="top" class="td_admin_2">
	     &nbsp;
	</td>
	<%end if%>
</tr>
<%
catinner.movenext
%>
<%if not catinner.eof then%>
<tr>
     <td align="right" colspan="6"  class="td_line_between">
     </td>
</tr>
<%end if%>
<%loop
%>
</table>
</td>
     <td colspan="2" class="td_admin_4_right_align">&nbsp;
     </td>
</tr>
<tr>
     <td align="right" colspan="7"  class="td_line_between">
     </td>
</tr>
<tr>
     <td align="right" colspan="7"  class="td_line_between">
     </td>
</tr>
<%	
catinner.close 
end if%>
<%end if%>
<%end if%>

<%
pr.movenext
loop
pr.close 

%>
<tr>
   <td align="right" colspan="7" class="td_admin_4_right_align" height="5">
   </td>
</tr>
</table>
</div>
</body>
<%end if
set worker= nothing
set con = nothing%>
</html>
