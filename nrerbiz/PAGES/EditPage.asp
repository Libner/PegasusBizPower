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
<title>CyberServe תבניות תוכן</title>
<meta charset="windows-1255">
<link href="../dynamic_style.css" rel="STYLESHEET" type="text/css">
<script language="JavaScript">
<!--                                    


function FormSubmitFun(pid,cid){
	if (CheckFields())
	document.f1.submit();
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


function ifFieldEmpty(fieldName){
	field = window.document.all(fieldName);
	if (field.value=='')
	{
		window.alert('חובה לבחור את התמונה');
		field.focus();
		return false;
	}
	return true;
}
function CheckDel(field,pageID) 
//For picture deleting
{
  if(confirm("?האם ברצונך למחוק את התמונה") == true )
  {
	window.document.form_upload.action="Aimgadd.asp?C=0&F="+field+"&pageID="+pageID;
	window.document.form_upload.submit();
	return true;
  }
  return false;
}
//-->
</script>
</head>
<body marginwidth="0" marginheight="0" topmargin="0" leftmargin="0" rightmargin="0" bgcolor="#F9F9F9">

<% 'on error resume next
sqlstring="SELECT * from workers where loginName='" & userName & "' and  password='"& password &"'"
set worker=con.Execute(sqlstring)
if worker.EOF then%>
<p><center><font color="red" size="4"><b>You are not authorized to enter staff zone !!!</b></font></center></p>
<%else
  session("admin.username")=username
  session("admin.password")=password
	isRightMenu=Request.querystring("isright")
	maincat=request("maincat")
    catid=Request("catid") 
    catCode=Request("catCode")
    pageId=Request("pageId")
    parent=request.querystring("parentID")
	innerparent=request("innerparent")
	subcat=request("subcat")	
	siteid=session("siteid")
  
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
		''Response.Write "<BR>"&sqlUpdt
		''Response.End 
		con.execute(sqlUpdt)  
		'Response.Redirect "publications.asp?pageid=" & pageid & "&catid=" & catid
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
		'Response.Redirect "publications.asp?pageid=" & pageid & "&catid=" & catid
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
	perBackground=pr("Page_Background")
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
%>
<table  width="100%" align="center" cellspacing="0" cellpadding="0" border=0>
 <%
  sql11 = "SELECT Site_Name FROM Sites WHERE Site_Id=" & SiteId
  set sites = con.execute(sql11)%>
 <tr><td class="a_title_big" width="100%" dir=rtl><center> מנגנון ניהול אתר "<%=sites("Site_Name")%>"</center></td></tr>
 <%  set sites = Nothing %>
  <tr><td bgcolor="#e6e6e6" height="1"></td></tr>
	<tr>
		<td width="100%" nowrap class="td_title_2">
		<table border="0" cellpadding="2" cellspacing="0">
			<tr>
				<td  width="50%" align=left>
					<a class = "button_admin_1" href="publications.asp?pageId=<%=pageId%>&innerparent=<%=pageId%>&subcat=<%=Request("subcat")%>&catId=<%=category%>">חזרה</a>
				</td>
				<td width="50%" align=right>
					<table border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td>
								<!--
								<a class="button_admin_1" href="../../template/PagePreview.asp?maincat=<%=maincat%>&catId=<%=category%>&PageId=<%=pageId%>&id_site=1&show=1">תצוגה מקדימה</a>
								-->
								<a class="button_admin_1" target="_blank" href="../../template/default.asp?maincat=<%=maincat%>&catId=<%=category%>&PageId=<%=pageId%>&id_site=<%=siteid%>&show=1">תצוגה מקדימה</a>
							</td> 
							<td width="5"></td>
							<td>
								<a class = "button_admin_1" href = "editPage.aspx?pageId=<%=pageId%>&parentID=<%=pageId%>&innerparent=<%=parentId%>&catId=<%=category%>&maincat=<%=maincat%>&siteid=<%=siteid%>" >עדכון הדף</a>
							</td>
							<td width="5"></td>
							<td>
								<a class="button_admin_e" href="" onclick="return false;">עדכון מאפייני הדף</a>
							</td> 
						</tr> 
					</table> 
				</td> 
							
			</tr>
		</table>
		</td>
	</tr>
	<tr><td bgcolor="#e6e6e6" height="1"></td></tr>
	<tr>
		<td align="right" valign="middle" class="title_admin_1" nowrap><span dir=rtl>&nbsp;קטגוריה&nbsp;"<%=parentname%>"</span></td>
	</tr>
	<tr><td bgcolor="#e6e6e6" height="1"></td></tr>
	<tr>
		<td align=center width=100% bgcolor="#e6e6e6">
		<form name="f1" ACTION="EditPage.asp?innerparent=<%=innerparent%>&subcat=<%=subcat%>&parentID=<%=parent%>&maincat=<%=maincat%>&pageId=<%=pageId%>&catId=<%=catid%>" METHOD="post" ID="Form2">
		<table class="table_admin_2" border="0" width="98%" cellspacing="0" cellpadding="0" align="left">		
		<tr><td height=10 colspan=2 nowrap class="td_admin_4"></td></tr>
		<tr>
			<td align="right" class="td_admin_4_right_align"><font face="Arial (Hebrew)"><textarea name="title" dir="rtl" rows="1" cols="60"><%=pertitle%></textarea></font></td>
			<td align="right" class="td_admin_4">כותרת הדף</td>			
		</tr>
		<tr>
			<td align="right" class="td_admin_4_right_align"><input type=text  name="pageURL" value="<%=pageURL%>" size=84></td>
			<td align="right" class="td_admin_4">כתובת של אתר אחר או לתיקיה פנימית URL</td>
		</tr>
		<tr>
			<td align="right" class="td_admin_4_right_align">
			('<b>http://</b>...' וא '<b>www.</b>...' - לינק חיצוני)&nbsp;<font color=red><b>*</b></font>
			</td>
			<td class =  "td_admin_4_right_align"></td>
		</tr>
		<tr>
			<td align="right" class="td_admin_4_right_align">&nbsp;&nbsp;%&nbsp;/&nbsp;px&nbsp;&nbsp;<input type=text  name="PageWidth" value="<%=Page_Width%>" size=5 maxlength=6></td>
			<td align="right" class="td_admin_4">רוחב דף</td>
		</tr>				
		<tr>
			<td align=center class="td_admin_4" colspan=2>
				<a  class="button_admin_e" href="" onClick="return FormSubmitFun('<%=pageid%>','<%=catid%>');">עדכן</a>
				<input type="hidden" name="pageId" value="<%=pageId%>">
				<input type="hidden" name="catCode" value="<%=catCode%>">
				<input type="hidden" name="parentID" value="<%=parentId%>">
				<input type="hidden" name="editP" value="1">
				<input type="hidden" name="place" value="<%=place%>">
				<input type="hidden" name="catid" value="<%=catid%>">
			</td>
		</tr>
		<tr><td height=10 colspan=2 nowrap class="td_admin_4"></td></tr>
		</table>
		</form>
		</td></tr>		
		<tr><td height="10" nowrap></td></tr>
		<tr><td height=1 colspan=2 bgcolor="#C0C0C0"></td></tr>		
		<tr><td colspan=2 bgcolor="#DBDBDB">
		<table cellpadding=0 cellspacing=0 width=100% border=0>		
		<tr><td height="10" nowrap></td></tr>
		<form ACTION="Aimgadd.asp?C=1&F=Page_background&pageID=<%=pageID%>" METHOD="post" ENCTYPE="multipart/form-data" ID="form_upload" name="form_upload">						
		<%If trim(perBackground) = "" Or IsNull(perBackground) = true then %>
		<tr>
			<td align="right" width="530" nowrap>&nbsp;</td>
			<td align="right" width="100" nowrap rowspan="2" valign=middle class="10normalB"><b>תמונה</b></td>
		</tr>
		<tr valign=middle>
			<td align="right" valign="middle">
			<table border="0" cellspacing="0" cellpadding="0" align="right">
				<tr valign=middle>        
				<td align="right" style="padding-right:2px;"><INPUT type="submit" class="form_button" onclick="return ifFieldEmpty('UploadFile2')" value="העלאת תמונה" id=submit1 name=submit1></td>
				<td align="right" style="padding-right:2px;"><INPUT TYPE="FILE" NAME="UploadFile2" size=30 ID="UploadFile2"></td>
				</tr>
			</table>       
			</td>
		</tr>
		<%Else%>
		<tr>
			<td align=right><img id="imgPict" name="imgPict" src="../../download/page_backgrounds/<%=perBackground%>" border="0" hspace=2 ></td>						
			<td align=right rowspan="2" nowrap valign=top class="10normalB"><b>תמונה</b></td>
		</tr>
		<tr><td colspan=2 height="10"></td></tr>
		<tr>
			<td align=right>
			<table border="0" cellspacing="0" cellpadding="0" align="right">
				<tr valign=middle>        
				<td align="right" style="padding-right:2px;"><INPUT type="button" class="form_button" ONCLICK="return CheckDel('delScreen','<%=pageID%>');" value="מחיקת תמונה" id=button1 name=button1></td>
				<td align="right" style="padding-right:2px;"><INPUT type="submit" class="form_button" onclick="return ifFieldEmpty('UploadFile2')" value="החלפת תמונה" id=submit2 name=submit2></td>
				<td align="right" style="padding-right:2px;"><INPUT TYPE="FILE" NAME="UploadFile2" size=30 ID="UploadFile2"></td>
				</tr>
			</table>
			</td>
		</tr>
		<%End If%>		
		</form>
		<tr><td height="10" nowrap></td></tr>
		</table></td></tr>		
		<tr><td height=1 colspan=2 bgcolor="#C0C0C0"></td></tr>
		<tr><td height="10" nowrap></td></tr>
</table>
<%end if%>
</body>
</html>
