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
<script language="JavaScript">
<!--                                    


function FormSubmitFun(pid,cid){
	if (CheckFields())
	document.f1.submit();
	//document.location.href="event.aspx?pageid=" + pid + "&catid=" + cid
return false;
}

function CheckFields() {
	if (document.f1.title.value==''){
		alert('חובה למלא את השדה טקסט');
		return false;
		}
	var val = parseInt(document.f1.PageWidth.value)
	if (isNaN(val) || val < 10 || val > 1600) {
		alert('רוחב דף שגוי');
		document.f1.PageWidth.select();
		return false;
		}	
	return true;
  }

function CheckDel() {
  return (confirm("?האם ברצונך למחוק את התמונה"))    
}

function onChange_balign(myValue){
  if (myValue!='')
     window.document.f1.bSpace.disabled=true;
  else
     window.document.f1.bSpace.disabled=false;
}

<!--End-->
//-->
</script>

<html>

<head>
<title>CyberServe תבניות תוכן</title>
<meta charset="windows-1255">
<link href="../dynamic_style.css" rel="STYLESHEET" type="text/css">
</head>
<body  class="body_admin">
<div align="right">

<% 'on error resume next
sqlstring="SELECT * from workers where loginName='" & userName & "' and  password='"& password &"'"
set worker=con.Execute(sqlstring)
if worker.EOF then%>
<p><center><font color="red" size="4"><b>You are not authorized to enter staff zone !!!</b></font></center>
<%else
  session("admin.username")=username
  session("admin.password")=password
  siteid=session("siteid")
isRightMenu=Request.querystring("isright")
    catid=Request("catid") 
    catCode=Request("catCode")
    pageId=Request("pageId")
    parent=request.querystring("parentID")
innerparent=request("innerparent")
subcat=request("subcat")	
'   Id=Request("Id") 
 'place=Request("place")
 'if place=nil then
	'place="0"
' end if	
' newPlace=CInt(place)+1
  
if Request.Form("homePage")="1" then
homeVis=1
else
	homeVis=0
end if	
%>
<%
		m=request("mm")
		d=request("dd")
		y=request("yy")
		if (y="" or isnull(y)) and not isnull(m) and m<>"" and not isnull(d) and d<>"" then
			y=Year(Date())
		end if
		'Response.Write m&"/"&d&"/"&y & " isdate="& isdate(m&"/"&d&"/"&y)
		if isdate(m&"/"&d&"/"&y) then
			'pdate="'"&dateserial(y,m,d)&"'"
			pdate="'"&m&"/"&d&"/"&y&"'"
		else
			pdate="NULL"
		end if

if request("editP")="1" then
     spaceUpd=Request.Form("bSpace")
	 vspaceUpd=Request.Form("bVSpace")
     if trim(Request.Form("balign"))<>"" then
	    alignUpd=Request.Form("balign")
	    spaceUpd="0"
	 else  
        alignUpd=""
	 end if
	 
	 if pageId<>nil then
		sqlUpdt="update pagesTav set Page_Title='"&Left(sFix(Request("title")),150)&"',Page_URL='"&sFix(Request("pageURL"))&"',Page_Date=" & pdate & ",Page_Width='" & Request("PageWidth") & "' ,Page_Background_align='"& alignUpd &"',Page_Background_Space='"& spaceUpd &"',Page_Background_VSpace='"& vspaceUpd &"' where Page_ID="&pageId&" "
		'Response.Write "<BR>"&sqlUpdt
		con.execute(sqlUpdt)  
		Response.Redirect "publications.asp?pageid=" & pageid & "&catid=" & catid
	 else
		set count=con.execute("select Count(Page_Order) AS Pcount from pagesTav where Category_Id="& catId & " and Page_Parent=" & parent)
		countpage=count("Pcount")
		if countpage>0 then
			ishome=0
		else
			ishome=1
		end if
        strng="update pagesTav  set Page_Order=Page_Order+1 WHERE Category_Id="&catid&" and Page_Parent="&parent
	    con.Execute(strng) 
	    sqlstring="INSERT INTO pagesTav(Category_Id,Page_Parent,Page_order,Page_Title,Page_Home,Page_URL,Page_Date,Page_Background_align,Page_Background_Space,Page_Background_VSpace,Page_Width) VALUES ("& catid &"," & parent &", 1 ,'"&Left(sFix(Request.form("title")),150)&"'," & ishome & ",'"&sFix(Request.form("pageURL"))& "'," & pdate & ",'"& alignUpd &"','"& spaceUpd &"','"& vspaceUpd &"','"& Request("PageWidth") &"') " 
	    con.Execute(sqlstring)
		set pag=con.execute("SELECT * from pagesTav ORDER BY Page_ID DESC")
	    pageId=CStr(pag("Page_ID"))
		pag.close
		Response.Redirect "publications.asp?pageid=" & pageid & "&catid=" & catid
    end if
    'rem: background ****************************
        'Response.Write("<br>bg="& Request.Form("isbackground"))
        if cStr(trim(Request.Form("isbackground")))=cStr("-1") then
            con.execute ("update pagesTav set Page_Background=null WHERE Page_Id=" & pageId)
        elseif Request.Form("isbackground")<>nil then
             backgroundUpd="Pict"& Request.Form("isbackground") &".jpg"
             con.execute ("update pagesTav set Page_Background='" & backgroundUpd &"' WHERE Page_Id=" & pageId)
        end if
        'rem: END background ****************************%>
<script>
    //window.location.href="event.aspx?innerparent=<%=innerparent%>&subcat=<%=subcat%>&pageId=<%=pageId%>"
</script>
<% end if%>

<%Page_Width=PageWidth
if not isnull(pageId) and pageid<>"" then
  sqlSt="SELECT * FROM pagesTav WHERE Page_Id="&pageId&" "
  set pr=con.Execute(sqlSt)
category=pr("Category_Id")
pertitle=pr("Page_Title")
Page_Width=pr("Page_Width")
pageURL=pr("Page_URL")
pageDate=pr("Page_Date")
perBackgroundNum=pr("Page_Background")
peralign=pr("Page_Background_align")
perSpace=pr("Page_Background_Space")
perVSpace=pr("Page_Background_VSpace")

if trim(perVSpace)="" OR isNull(perVSpace) then
   perVSpace="0"
end if
if trim(perSpace)="" OR  isNull(perSpace) then
   perSpace="0"
end if
if not isnull(pageDate) and pageDate<>"" then 
Dayp=Day(pageDate)
Monthp=Month(pageDate)
Yearp=Year(pageDate)
end if

	if parent<>"" and not isNull(parent)  and parent<>"0" then
		sqlcat=" select Page_Title from pagesTav where Page_ID=" &parent
		set par=con.Execute(sqlcat)
		parentname=par("Page_Title")
	else
		sqlcat="Select * from Publication_Categories where Publication_Category_ID="&category&" "
		set cat=con.Execute(sqlcat)
		parentname=cat("Publication_Category_Name")
	end if
else
	if parent<>"" and not isNull(parent) and parent<>"0" then
		sqlcat=" select Page_Title,Page_URL from pagesTav where Page_ID=" &parent
		set par=con.Execute(sqlcat)
		parentname=par("Page_Title")
		pageURL=par("Page_URL")
	else
		sqlcat="Select * from Publication_Categories where Publication_Category_ID="&catid&" "
		set cat=con.Execute(sqlcat)
		parentname=cat("Publication_Category_Name")
	end if
end if

'rem: background ****************************
'Response.Write ("<br>peralign="&peralign)
if trim(peralign)<>"" then
   str_perSpace="disabled"
end if
strCheckDefault="checked"
strCheck0=""
strCheck1=""
strCheck2=""
strCheck3=""
strCheck4=""

if perBackgroundNum="Pict0.jpg" then
   strCheck0="checked"
   strCheckDefault=""
elseif perBackgroundNum="Pict1.jpg" then
   strCheck1="checked"
   strCheckDefault=""
elseif perBackgroundNum="Pict2.jpg" then
   strCheck2="checked"
   strCheckDefault=""
elseif perBackgroundNum="Pict3.jpg" then
   strCheck3="checked"      
   strCheckDefault=""
end if
'rem: END background ****************************
%>
<table class="table_admin_2" border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
 <tr>
    <td class="a_title_big" width="100%" ><%if pageId=nil then%>הוספת דף<%else%>"<%=pertitle%>" עדכון דף<%end if%>
    </td>
  </tr>
       <td class="td_admin_4" align="center" valign="middle" height="2">
       </td>
  </tr>
  <tr>
    <td align="right" valign="middle" class="td_title_2" nowrap><span dir=rtl>&nbsp;קטגוריה&nbsp;"<%=parentname%>"</span></td>
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
					<a  class="button_admin_e" href="" onClick="return FormSubmitFun('<%=pageid%>','<%=catid%>');">שלח</a>
				</td>
	<%if catid<>nil then  %>
				<td align="left" valign="middle" nowrap>
					<a  class="button_admin_1" href="publications.asp?innerparent=<%=innerparent%>&subcat=<%=subcat%>&catId=<%=catId%>">
						חזרה לניהול תבניות
					</a>
				</td>
	<%end if%>
				<%if pageId<>nil then%>		
				<td align="left" valign="middle" nowrap>
					<a  class="button_admin_1" href="event.aspx?innerparent=<%=innerparent%>&subcat=<%=subcat%>&pageId=<%=pageId%>&catId=<%=catId%>">
						חזרה לניהול הדף
					</a>
				</td>
				<%end if%>
				<td width="100%">&nbsp;</td>
			</tr>
		</table>
	</td>
</tr>
  <tr>
  <tr>
       <td class="td_admin_4" align="center" valign="middle" nowrap>&nbsp;
       </td>
  </tr>
</table>
<table class="table_admin_2" width="100%" border="0"  cellpadding="2" cellspacing="0">



<form name="f1" ACTION="AddPage.asp?innerparent=<%=innerparent%>&subcat=<%=subcat%>&parentID=<%=parent%>" METHOD="post">
   <tr>
         <td align="right" class =  "td_admin_4_right_align"><font face="Arial (Hebrew)"><textarea name="title" dir="rtl" rows="1" cols="60"><%=pertitle%></textarea></font></td>
         <td align="right" class =  "td_admin_4">כותרת הדף</td>
         </td>
    </tr>
<tr>
	<td align="right" class =  "td_admin_4_right_align"><input type=text  name="pageURL" value="<%=pageURL%>" size=84></td>
	<td align="right" class =  "td_admin_4">כתובת של אתר אחר או לתיקיה פנימית URL</td>
</tr>
<tr>
	<td align="right" class =  "td_admin_4_right_align">
	('<b>http://</b>...' וא '<b>www.</b>...' - לינק חיצוני)&nbsp;<font color=red><b>*</b></font>
	</td>
	<td class =  "td_admin_4_right_align"></td>
</tr>

<tr>
	<td align="right" class =  "td_admin_4_right_align">px&nbsp;<input type=text  name="PageWidth" value="<%=Page_Width%>" size=5 maxlength=4></td>
	<td align="right" class =  "td_admin_4">רוחב דף</td>
</tr>				

<input type="hidden" name="pageId" value="<%=pageId%>">
<input type="hidden" name="catCode" value="<%=catCode%>">
<input type="hidden" name="parentID" value="<%=parentId%>">
<input type="hidden" name="editP" value="1">
<input type="hidden" name="place" value="<%=place%>">
<input type="hidden" name="catid" value="<%=catid%>">


<%if false then%>
<%'rem: background ****************************%>
<tr><td height="1" bgcolor="navy"></td></tr>

<tr>
  <td align="right" width="100%">
    <table align="right" border="0" cellspacing="0" cellpadding="0" width="100%">
<tr>
    <td align="right" width="85%"  bgcolor="#feffe7">
    <table align="right" bgcolor="#feffe7" border="0" cellspacing="1" cellpadding="0" width="100%">
			<tr><td align="center" bgcolor="navy"><font color="#ffffff"><strong>רושיי</strong></font></td></tr>
			<tr><td height="1" bgcolor="navy"></td></tr>
			<tr><td height="5"></td></tr>
			<tr>
   		    <td align="right" width="100%"  bgcolor="#feffe7">
            <table align="right" bgcolor="#feffe7" border="0" cellspacing="1" cellpadding="0" width="100%">
			<tr>
			    <td width="50%" align="center" bgcolor="#feffe7" nowrap>
				   <table align="right" bgcolor="#feffe7" border="0" cellspacing="1" cellpadding="0" width="100%">
				      <tr>
				        <td align="right" bgcolor="#feffe7" nowrap>
				           <%i=0%>
				           <select name="bVSpace" size="1">
					         <%do while i<=200%>
					           <option value="<%=i%>" <%if cStr(perVSpace)=cStr(i)  then%>selected<%end if%>><%=i%></option>
				               <%i=i+1
				             Loop%>
				           </select>&nbsp;:&nbsp;הלעמלמ חוור
				         </td>
			          </tr>
				      <tr>			
		                <td align="right" width="100%" bgcolor="#feffe7" nowrap>
				          <%i=0%>
				         <select name="bSpace" size="1" <%=str_perSpace%>> 
					     <%do while i<=200%>
					        <option value="<%=i%>" <%if cStr(perSpace)=cStr(i) then%>selected<%end if%>><%=i%></option>
				            <%i=i+1
				         Loop%> 
				         </select>&nbsp;: &nbsp;הצקהמ חוור
				        </td>
			          </tr>
				   </table> 
				</td>
				<td width="50%" align="center" bgcolor="#feffe7" nowrap>
				    <font face="Arial (Hebrew)"><div dir="rtl">
				  	<select name="balign" size="4" onChange="return onChange_balign(this.options[this.selectedIndex].value);"> 
					    <option value="center" <%if peralign="center" then%>selected<%end if%>><%=reverse("זכרוממ")%></option>
						<option value="left" <%if peralign="left" then%>selected<%end if%>><%=reverse("לאמשל")%></option>
    					<option value="right" <%if peralign="right" then%>selected<%end if%>><%=reverse("ןימיל")%></option>
					    <option value="" <%if trim(peralign)="" then%>selected<%end if%>>רחא</option>
					</select>
					</div></font>
				</td>
			</tr>
			</table></td></tr>
			<tr>
				<td bgcolor="#feffe7" height="5"><table><tr><td></td></tr></table></td>
			</tr>
		</table>	
    </td>
    <td align="right" width="15%" bgcolor="#feffe7"  valign="top" nowrap>
       <table align="right" bgcolor="#feffe7" border="0" cellspacing="0" cellpadding="0" width="100%">
           <tr><td align="right" width="100%" bgcolor="#feffe7">&nbsp;<b>Background</b></td></tr> 
           <tr><td height="2"></td></tr>
           <tr><td height="1" bgcolor="navy"></td></tr>
       </table>
    </td>      
</tr></table></td>
</tr>
<tr><td height="5" bgcolor="#feffe7"></td></tr>
<tr>
	<td colspan="4" align="left" bgcolor="#ffffff">
	   <table align="center" border="0" bordercolor="#226699" cellpadding="0" cellspacing="0" width="100%">
        <tr>
	      <td width="100%" align="left" bgcolor="#226699">	    
	      <table align="center" border="0" cellpadding="0" cellspacing="1" width="100%">
	      <!--form method="post" name="ff" action="event.aspx?pageId=<%=pageId%>&catId=<%=catId%>"-->  
	      <tr>
	          <td bgcolor="<%if cStr(strCheck0)<>cStr("checked") then%>#ffffff<%end if%>" align="center"><input <%=strCheck0%> type="radio" name="isbackground" value="0"></td>
	          <td bgcolor="<%if cStr(strCheck1)<>cStr("checked") then%>#ffffff<%end if%>" align="center"><input <%=strCheck1%> type="radio" name="isbackground" value="1"></td>
	          <td bgcolor="<%if cStr(strCheck2)<>cStr("checked") then%>#ffffff<%end if%>" align="center"><input <%=strCheck2%> type="radio" name="isbackground" value="2"></td>
	          <td bgcolor="<%if cStr(strCheck3)<>cStr("checked") then%>#ffffff<%end if%>" align="center"><input <%=strCheck3%> type="radio" name="isbackground" value="3"></td>
	          <td bgcolor="<%if cStr(strCheckDefault)<>cStr("checked") then%>#ffffff<%end if%>" align="center"><input <%=strCheckdefault%> type="radio" name="isbackground" value="-1"></td>
	      </tr>
	      <tr>
	         <%for j=0 to 3%>
	             <td valign="top" width="170" align="left" valign="middle" bgcolor="#e5e5e5"><img src="../../backgrounds/Pict<%=j%>.jpg" vspace="10"></td>
	         <%next%>
	         <td bgcolor="#e5e5e5" width="100" align="center" valign="top"><font size="3" color="red"><br><b>אלל</b></font></td>
	      </tr>
	      <!--/form-->
	     </table></td></tr>
	   </table>
	</td>
</tr>
<%end if
'rem: END background ****************************%>



</form>


</table>
<%end if%>
</div>
</body>

</html>
