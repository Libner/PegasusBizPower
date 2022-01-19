
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<%
cod = trim(Request("cod"))
arch = trim(Request("arch"))
if arch = "" then
	arch = 0
end if
if arch = "0" then
	set_arch = 1
else
	set_arch = 0
end if

where_arch = " and appeal_deleted=" & arch

dim where_app(5)
where_app(1) = " and appeal_status = '1'"
where_app(2) = " and appeal_status = '2'"
where_app(3) = " and appeal_status = '3'"
where_app(4) = ""

if (Request("prodId") <> nil) then 'and Request.QueryString("delProd") = nil) then
	prodId = Request("prodId")
	where_product_id = " and questions_id=" & prodId
	set prod = con.GetRecordSet("Select product_name from products where product_id=" & prodId)
	if not prod.eof then
		title_prodName= prod("product_name")
	end if
	set prod = nothing
	title_ = "ניהול משובים מדיוור"
else
	where_product_id = ""
	title_ = "ניהול משובים מדיוור"
end if

if (Request("sekerId") <> nil) then 'and Request.QueryString("delProd") = nil) then
	sekerId = Request("sekerId")
	where_seker_id = " and product_id=" & sekerId
else
	where_seker_id = ""	
end if

if Request.QueryString("clientId") <> nil then
	clientId = Request.QueryString("clientId")
	where_client_id = " and PEOPLE_ID=" & clientId
else
	where_client_id = ""	
end if

if trim(Request("orgname")) <> "" then
    orgname = trim(Request("orgname")) 
	where_orgname = " and LOWER(PEOPLE_COMPANY) LIKE '%"& LCase(trim(Request.Form("orgname"))) &"%'"		
else
	where_orgname = ""
end if
if trim(Request("clname")) <> "" then
    clname = trim(Request("clname"))
	where_clname = " and LOWER(PEOPLE_name) LIKE '%"& LCase(trim(Request.Form("clname"))) &"%'"		
else
	where_clname = ""
end if
	
dim text_no(5)	
text_no(1) = "לא נמצאו משובים מדיוור חדשים"
text_no(2) = "לא נמצאו משובים מדיוור בטיפול"
text_no(3) = "לא נמצאו משובים מדיוור סגורים"
text_no(4) = "לא נמצאו משובים מדיוור"

dim sortby(16)	
sortby(1) = "CAST(appeal_date as float), appeal_id DESC"
sortby(2) = "CAST(appeal_date as float) DESC, appeal_id DESC"
sortby(3) = "appeal_id"
sortby(4) = "appeal_id DESC"
sortby(5) = "PEOPLE_EMAIL, CAST(appeal_date as float) DESC, appeal_id DESC"
sortby(6) = "PEOPLE_EMAIL DESC, CAST(appeal_date as float) DESC, appeal_id DESC"
sortby(7) = "product_name, CAST(appeal_date as float) DESC, appeal_id DESC"
sortby(8) = "product_name DESC, CAST(appeal_date as float) DESC, appeal_id DESC"
sortby(9) = "PEOPLE_NAME, CAST(appeal_date as float) DESC, appeal_id DESC"
sortby(10) = "PEOPLE_NAME DESC, CAST(appeal_date as float) DESC, appeal_id DESC"
sortby(11) = "appeal_status, CAST(appeal_date as float) DESC, appeal_id DESC"
sortby(12) = "appeal_status DESC, CAST(appeal_date as float) DESC, appeal_id DESC"
sortby(11) = "PEOPLE_COMPANY, CAST(appeal_date as float) DESC, appeal_id DESC"
sortby(12) = "PEOPLE_COMPANY DESC, CAST(appeal_date as float) DESC, appeal_id DESC"

sort = Request("sort")	
if sort = nil then
	sort = 2
end if

' ------ start deleting appeals ----
if Request.QueryString("delProd") <> "" then
	delappid = Request.QueryString("appId")	
	'------ start deleting the new message from XML file ------
	xmlFilePath = "../../../download/xml_forms/bizpower_forms.xml"
	set objDOM = Server.CreateObject("Microsoft.XMLDOM")
	objDom.async = false			
	if objDOM.load(server.MapPath(xmlFilePath)) then
		set objNodes = objDOM.getElementsByTagName("FORM")
		for j=0 to objNodes.length-1
			set objTask = objNodes.item(j)
			node_app_id = objTask.attributes.getNamedItem("APPEAL_ID").text										
			if trim(delappid) = trim(node_app_id) Then					
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
	con.ExecuteQuery("Delete from Form_Value where appeal_id =" & delappid )
	con.ExecuteQuery("Delete from appeals where appeal_id =" & delappid)
	Response.Redirect "feedbacks.asp?prodId="&Request.QueryString("quest_id")
end if 

' ------ end deleting appeals --------
' ------ transfering to archive --------

if trim(Request.Form("trapp")) <> "" then
	if Request.Cookies("bizpegasus")("Archive_appeal")="1" then
			sqlstr = "UPDATE appeals SET appeal_deleted=" & set_arch & " Where appeal_id in (" & Request.Form("trapp") & ")"
			'Response.Write(sqlstr)
			con.ExecuteQuery(sqlstr)
	end if
end if

 if trim(lang_id) = "1" then
        arr_Status = Array("","חדש","בטיפול","סגור")	
 else
        arr_Status = Array("","new","active","close")	
 end if

%>	
<%   
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 49 Order By word_id"				
	  set rstitle = con.getRecordSet(sqlstr)		
	  If not rstitle.eof Then
	    dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(arr_title(0)) = arr_title(1)
			End If
		Next
	  End If
	  set rstitle = Nothing	
%> 	
<script LANGUAGE="javascript">
<!--
var oPopup = window.createPopup();

function ProdDropDown(obj)
{
    oPopup.document.body.innerHTML = Prod_Popup.innerHTML; 
    oPopup.show(obj.offsetWidth-190, obj.offsetHeight, 190, 200, obj);
    
}

function StatusDropDown(obj)
{
    oPopup.document.body.innerHTML = Status_Popup.innerHTML; 
    oPopup.show(-54, 19, 70, 90, obj);
    
}
function CheckDelProd() {
  <%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק את המשוב"
     Else
		str_confirm = "Are you sure want to delete the feedback?"
     End If   
  %>
 return (confirm("<%=str_confirm%>")) 
}

function cball_onclick() {
	var strid = new String(document.form1.ids.value);
	var arrid = strid.split(',');
	for (i=0;i<arrid.length;i++)
		document.forms('form1').elements('cb'+ arrid[i]).checked = document.form1.cball.checked ;
	
}
function checktransf(){
		var fl = 0;
		document.form1.trapp.value = '';
		if(document.form1.ids)
		{
			var strid = new String(document.form1.ids.value);
			if(strid != "")
			{
				var arrid = strid.split(',');
				for (i=0;i<arrid.length;i++){
					if (document.forms('form1').elements('cb'+ arrid[i]).checked)
					{	document.form1.trapp.value = document.form1.trapp.value + arrid[i] + ',';
						fl = 1;
					}	
				}
		    <%
			If trim(lang_id) = "1" Then
				str_confirm = "? האם ברצונך להעביר את הטפסים המסומנים"
			Else
				str_confirm = "Are you sure want to transfer the selected feedbacks?"
			End If   
		    %>
			if (fl && confirm("<%=str_confirm%>")){
					var txtnew = new String(document.form1.trapp.value);
					document.form1.trapp.value = txtnew.substr(0,txtnew.length - 1);
					return true;
				}
				else if (fl) return false;
			}
		  <%
			If trim(lang_id) = "1" Then
				str_confirm = "! נא לסמן הטפסים להעברה"
			Else
				str_confirm = "Please select feedbacks to transfer !"
			End If   
		  %>			
		window.alert("<%=str_confirm%>");
		return false;	
		}	
		
		return false;
	}
var oClient = null;	
function showcard(Id,gID)
  {
    strOpen="../groups_clients/addpeople.asp?people_Id="+Id+"&groupId="+gID;
	oClient = window.open(strOpen, "Client", "scrollbars=1,toolbar=0,top=50,left=150,width=450,height=350,align=center,resizable=0");
	oClient.focus();
	return false;
  }

var prod;
function openPreview(prodId)
{
	prod = window.open("../products/check_form.asp?prodId="+prodId,"Product","toolbar=0, menubar=0, resizable=1, scrollbars=1, fullscreen = 0, width=700, height=600, left=10, top=20");
	if((prod.document != null) && (prod.document.body != null))
	{
		prod.document.title = "Product Preview";
		prod.document.body.style.margintop  = 20;
	}	
	return false;
}

//-->
</script>
<body>
<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">
<tr><td width=100% colspan=2>
<!-- #include file="../../logo_top.asp" -->
</td></tr>
<%numOftab = 2%>
<%numOfLink = 3%>
<%topLevel2 = 21 'current bar ID in top submenu - added 03/10/2019%>
<tr><td width=100% colspan=2>
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title" colspan=2>
 <SELECT dir="<%=dir_obj_var%>" size=1 ID=prodId name=prodId class="sel" style="width:310px;font-size:10pt" onChange="document.location.href='feedbacks.asp?arch=<%=arch%>&sort=<%=sort%>&prodID='+this.value">	
 <OPTION value="" id=word3 name=word3><!--כל הטפסים--><%=arrTitles(3)%></OPTION>
<%
	sqlstr = "Select product_id, Replace(product_name, Char(59), '&#59;')  FROM Products Where FORM_MAIL = '1' And "&_
	" product_number = '0' AND ORGANIZATION_ID=" & OrgID & " order by product_name"
	'Response.Write sqlstr
	'Response.End
	set rs_products = con.GetRecordSet(sqlstr)
	if not rs_products.eof then 		
		arr_products = rs_products.getRows()
	end if
	set rs_products=nothing	
	
	If IsArray(arr_products) Then
		If trim(projectID) = "" And trim(companyID) = "" Then
			from = "comp_list"
		Else
			from = "pop_up"
		End If	 
		
		For i=0 To  Ubound(arr_products,2) 				
				prod_Id = trim(arr_products(0,i))
				product_name = trim(arr_products(1,i))%>
	<OPTION value="<%=prod_Id%>" <%If trim(prodId) = trim(prod_Id) Then%> selected <%End If%>> <%=product_name%> </OPTION>
<%	Next	
 End If%>	</SELECT></td></tr>		    		       	       
 <tr><td valign=top>
 <table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">  
 <tr>    
    <td width="100%" valign="top" align="center">
	<table width="100%" cellspacing="0" cellpadding="0" align=center border="0" bgcolor="#ffffff" dir="<%=dir_var%>">
    <form action="feedbacks.asp?arch=<%=arch%>&clientId=<%=Request("clientId")%>&prodId=<%=Request("prodId")%>&cod=<%=Request("cod")%>&sort=<%=Request("sort")%>" method="POST" id="form1" name="form1" onsubmit="return checktransf();">
	<tr>
	<td width="100%" align="center" valign=top>

<!-- start code --> 
<%	
	urlSort = "feedbacks.asp?arch=" & trim(arch) & "&prodId=" & trim(prodId) & "&cod=" & trim(cod) & "&clname=" & Server.URLEncode(clname) & "&orgname=" & Server.URLEncode(orgname) & "&sekerId=" & Server.URLEncode(sekerId)
    sqlstr = "EXECUTE get_feedbacks '" & sFix(orgname) & "','" & sFix(clname) & "','','" & cod & "','" & OrgID & "','" & sortby(sort) & "','','','','','','','" & prodId & "','','" & arch & "','" & sFix(usname) & "','" & sekerId & "'"	
    'Response.Write sqlStr
	set app=con.GetRecordSet(sqlStr)
	app_count = app.RecordCount
	if Request("Page")<>"" then
		Page=request("Page")
	else
		Page=1
	end if
	%>
	<input type="hidden" name="trapp" value="" ID="trapp">		
	<%	if not app.eof then %>
	<table BORDER="0" CELLSPACING="1" CELLPADDING="1">	
	<tr>
		<td width="40" nowrap class="title_sort" align=center><span id=word4 name=word4><!--מחיקה--><%=arrTitles(4)%></span></td>
		<td width="40" nowrap class="title_sort" align=center><span id="word5" name=word5><!--עדכון--><%=arrTitles(5)%></span></td>
		<td width="55" nowrap dir="<%=dir_obj_var%>" align=center class="title_sort<%if sort=1 or sort=2 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=2 then%>1<%elseif sort=1 then%>2<%else%>2<%end if%>" target="_self"><span id="word6" name=word6><!--התקבל--><%=arrTitles(6)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort)="1" then%>down<%elseif trim(sort)="2" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="0"></a></td>		
		<td width="120" nowrap dir="<%=dir_obj_var%>" align="<%=align_var%>" class="title_sort<%if sort=5 or sort=6 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=6 then%>5<%elseif sort=5 then%>6<%else%>6<%end if%>" target="_self">&nbsp;Email&nbsp;<img src="../../images/arrow_<%if trim(sort)="5" then%>down<%elseif trim(sort)="6" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	
		<td width="100" nowrap dir="<%=dir_obj_var%>" align="<%=align_var%>" class="title_sort<%if sort=9 or sort=10 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=10 then%>9<%elseif sort=9 then%>10<%else%>10<%end if%>" target="_self">&nbsp;<span id="word7" name=word7><!--נמען--><%=arrTitles(7)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort)="9" then%>down<%elseif trim(sort)="10" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td width="100" nowrap dir="<%=dir_obj_var%>" align="<%=align_var%>" class="title_sort<%if sort=11 or sort=12 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=12 then%>11<%elseif sort=11 then%>12<%else%>12<%end if%>" target="_self">&nbsp;<span id="word8" name=word8><!--חברה--><%=arrTitles(8)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort)="11" then%>down<%elseif trim(sort)="12" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td width="100%" dir="<%=dir_obj_var%>" id=td_prod name=td_prod align="<%=align_var%>" class="title_sort<%if sort=7 or sort=8 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=7 then%>8<%elseif sort=8 then%>7<%else%>7<%end if%>" target="_self">&nbsp;<span id="word9" name=word9><!--סוג טופס--><%=arrTitles(9)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort)="7" then%>down<%elseif trim(sort)="8" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td class="title_sort">&nbsp;</td>		
		<td width="40" nowrap dir="<%=dir_obj_var%>" align=center class="title_sort<%if sort=3 or sort=4 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=4 then%>3<%elseif sort=3 then%>4<%else%>4<%end if%>" target="_self">&nbsp;ID&nbsp;<img src="../../images/arrow_<%if trim(sort)="3" then%>down<%elseif trim(sort)="4" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td width="48" nowrap dir="<%=dir_obj_var%>" class="title_sort" align="<%=align_var%>">&nbsp;<span id="word10" name=word10><!--'סט--><%=arrTitles(10)%></span>&nbsp;<IMG style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 ALT="בחר מרשימה" align=absmiddle onmousedown="StatusDropDown(this)"></td>
		<td width="20" nowrap  dir="<%=dir_obj_var%>" class="title_sort" align=center><%if not app.eof then%><INPUT type="checkbox" id=cball name=cball LANGUAGE="javascript" onclick="return cball_onclick()" title="העברה לארכיון"><%end if%></td>
	</tr>
	<%	
		If trim(RowsInList) <> "" AND IsNumeric(RowsInList) = true Then
			 PageSize = RowsInList
		Else	
     		 PageSize = 10
		End If		
		app.PageSize = PageSize
		app.AbsolutePage=Page
		recCount=app.RecordCount 		
		NumberOfPages = app.PageCount
		i=1
		j=0
		ids = "" 'list of appeal_id
		do while (not app.eof and j<app.PageSize)
		appid=app("appeal_id")	
		PEOPLE_ID = app("PEOPLE_ID")		
		PEOPLE_EMAIL = app("PEOPLE_EMAIL")
		PEOPLE_COMPANY = app("PEOPLE_COMPANY")		
		If trim(PEOPLE_COMPANY) = "" Or IsNull(PEOPLE_COMPANY) Then
			If trim(lang_id) = "1" Then
			PEOPLE_COMPANY = "אין"
			Else
			PEOPLE_COMPANY = "No"
			End If
		End If	
		PEOPLE_NAME = app("PEOPLE_NAME")
		If trim(PEOPLE_NAME) = "" Or IsNull(PEOPLE_NAME) Then
			If trim(lang_id) = "1" Then
			PEOPLE_NAME = "אין"
			Else
			PEOPLE_NAME = "No"
			End If
		End If	
		
		ids = ids & appid 
		groupID = app("groupID")		
		prod_id = app("product_id")
		quest_id = trim(app("questions_id"))
		If trim(quest_id) <> "" Then
			sqlstr = "Select PRODUCT_NAME FROM PRODUCTS WHERE PRODUCT_ID = " & quest_id
			set rs_name = con.getRecordSet(sqlstr)
			If not rs_name.eof Then
				if len(rs_name("product_name")) > 30 then
					product_name = left(rs_name("product_name"),27) & "..."
				else
					product_name = rs_name("product_name")
				end if	
			End if
			set rs_name = Nothing
		End If	
		appeal_status = trim(app("appeal_status"))	
		%>
		<tr>
			<td align="center" class="card" nowrap><a href="feedbacks.asp?arch=<%=arch%>&page=<%=page%>&quest_id=<%=quest_id%>&prodId=<%=prod_id%>&appid=<%=appid%>&delProd=1&cod=<%=cod%>&sort=<%=sort%>" ONCLICK="return CheckDelProd()"><IMG SRC="../../images/delete_icon.gif" BORDER=0 alt="מחיקת תשובה לסקר"></a></td>
			<td align="center" class="card" nowrap>&nbsp;<a href="feedback.asp?quest_id=<%=quest_id%>&prodId=<%=prod_id%>&appid=<%=appid%>" target="_self"><IMG SRC="../../images/edit_icon.gif" BORDER=0 alt="עדכון משובים מדיוור"></a>&nbsp;</td>
			<td class="card" align=center><a class="link_categ" HREF="feedback.asp?quest_id=<%=quest_id%>&prodId=<%=prod_id%>&appid=<%=appid%>" target="_self"><%=day(app("appeal_date"))%>/<%=month(app("appeal_date"))%>/<%=mid(year(app("appeal_date")),3,2)%></a></td>			
			<td class="card" nowrap align="<%=align_var%>"><a class="link_categ" HREF="feedback.asp?quest_id=<%=quest_id%>&prodId=<%=prod_id%>&appid=<%=appid%>" target="_self" style="font-size:11px">&nbsp;<%=PEOPLE_EMAIL%>&nbsp;</a></td>
			<td class="card" nowrap align="<%=align_var%>"><a class="link_categ" href="#" onclick="return showcard('<%=PEOPLE_ID%>','<%=groupID%>')" target="_self">&nbsp;<%=PEOPLE_NAME%>&nbsp;</a></td>
			<td class="card" nowrap align="<%=align_var%>"><a class="link_categ" HREF="feedback.asp?quest_id=<%=quest_id%>&prodId=<%=prod_id%>&appid=<%=appid%>" target="_self">&nbsp;<%=PEOPLE_COMPANY%>&nbsp;</a></td>
			<td class="card" nowrap align="<%=align_var%>"><a class="link_categ" href="#" OnClick="return openPreview('<%=quest_id%>')">&nbsp;<%=product_name%>&nbsp;</a></td>
			<td class="card" align="<%=align_var%>">
			<%key_table_width=150%>
		    <!--#include file="key_fields_f.asp"-->
			</td>			
			<td class="card" nowrap align=center><%=appid%></td>
			<td class="card" nowrap align=center><a class=status_num<%=appeal_status%> href="feedback.asp?quest_id=<%=quest_id%>&appid=<%=appid%>"><%=arr_Status(appeal_status)%></a></td>
			<td class="card" nowrap align=center><INPUT type="checkbox" id=cb<%=appid%> name=cb<%=appid%> title="העברה לארכיון"></td>
		</tr>
<%		app.movenext
		j=j+1
		if not app.eof and j <> app.PageSize then
		ids = ids & ","
		end if
		loop 	
		%>		
		</table>		
		<input type="hidden" name="ids" value="<%=ids%>" ID="ids">		
		</td>		
		</tr>
		<%End If%>
		</form>
		<% if NumberOfPages > 1 then%>
		<tr class="card">
		<td width="100%" align=center nowrap class="card">
			<table border="0" cellspacing="0" cellpadding="2">               
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
	           If Request.QueryString("numOfRow") <> nil Then
	               numOfRow = Request.QueryString("numOfRow")
	           Else numOfRow = 1
	           End If
	         %>
	         
	         <tr>
	         <%if numOfRow <> 1 then%> 
			 <td valign="middle" class="card"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=10*(numOfRow-1)-9%>&numOfRow=<%=numOfRow-1%>" title="לדפים הקודמים"><b><<</b></a></td>			                
			  <%end if%>
	             <td><font size="2" color="#060165">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	                  if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		                 <td align="center"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
	                  <%else%>
	                     <td align="center"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=i+10*(numOfRow-1)%>&numOfRow=<%=numOfRow%>"><%=i+10*(numOfRow-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td><font size="2" color="#060165">]</font></td>
				<%if NumberOfPages > cint(num * numOfRow) then%>  
					<td valign="middle"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=10*(numOfRow) + 1%>&numOfRow=<%=numOfRow+1%>" title="לדפים הבאים"><b>>></b></a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>			
	<tr><td align="center" height="20" class="card"><font style="color:#6E6DA6;font-weight:600"><!--נמצאו--><%=arrTitles(20)%>&nbsp; <%=app_count%> <span id=word11 name=word11>משובים מדיוור</span></font></td></tr>								
	<%End If%>
	<%if app.recordCount = 0 Then%>
	<tr>
		<td colspan=8 align="center" class="title_sort1" dir="<%=dir_obj_var%>"><!--לא נמצאו משובים מדיוור--><%=arrTitles(12)%></td>
	</tr>	
<%	end if ' if not app.eof%>			
</table></td>	
<td width=100 nowrap align="<%=align_var%>" valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width=100 border=0>
<tr><td align="<%=align_var%>" colspan=2 class="title_search" height=21><span id=word13 name=word13><!--חיפוש--><%=arrTitles(13)%></span></td></tr>
<FORM action="feedbacks.asp?sort=<%=sort%>&prodId=<%=prodId%>" method=POST id=form_search_comp name=form_search_comp target="_self">   
<tr><td colspan=2 align="<%=align_var%>" class="right_bar"><!--חברה--><%=arrTitles(14)%></td></tr>
<tr>
<td align="<%=align_var%>"><input type="image" onclick="form_search_comp.submit();" src="../../images/search.gif"></td>
<td align="<%=align_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:80;" value="<%=vFix(orgname)%>" name="orgname" ID="orgname">
</td></tr>
</FORM>
<FORM action="feedbacks.asp?sort=<%=sort%>&prodId=<%=prodId%>" method=POST id="form_search_contact" name="form_search_contact" target="_self">   
<tr><td colspan=2 align="<%=align_var%>" class="right_bar"><!--נמען--><%=arrTitles(21)%></td></tr>
<tr><td align="<%=align_var%>"><input type="image" onclick="form_search_contact.submit();" src="../../images/search.gif"></td>
<td align="<%=align_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:80;" value="<%=vFix(clname)%>" name="clname" ID="clname"></td></tr>
</FORM>
<tr><td colspan=2 height=10 nowrap></td></tr>
<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:95;line-height:110%;padding:3px"  href='feedbacks.asp?arch=<%if arch = "1" then%>0<%else%>1<%end if%>'><%if arch = "0" then%><span id=word15 name=word15><!--ארכיון--><%=arrTitles(15)%></span><%else%><span id="word16" name=word16><!--חזרה למושבים--><%=arrTitles(16)%></span><%end if%></a></td></tr>
<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:95;line-height:110%;padding:3px"  href='#' onclick="if (checktransf()) {document.form1.submit()}"><%if arch="0" then%><span id="word17" name=word17><!--העברה לארכיון--><%=arrTitles(17)%></span><%else%><span id="word18" name=word18><!--לטיפול חוזר--><%=arrTitles(18)%></span><%end if%></a></td></tr>
<tr><td align="<%=align_var%>" colspan=2 height="10" nowrap></td></tr>
</table></td></tr>
</table></td></tr>
</table>
</body>
<DIV ID="Status_Popup" STYLE="display:none;">
<div dir="<%=dir_obj_var%>" style="position:absolute; top:0; left:0; height:90; border:1px solid black; background-color: #D3D3D3" >	
<%For i=1 To Ubound(arr_Status)%>
	<DIV onmouseover="this.style.background='#9CCDF6'" onmouseout="this.style.background='#d3d3d3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#d3d3d3; border-bottom:1px solid black; padding:3px; cursor:hand"
	ONCLICK="parent.location.href='feedbacks.asp?prodId=<%=prodId%>&sort=<%=sort%>&cod=<%=i%>'">
    <%=arr_Status(i)%>
    </DIV>
<%Next%>    
   <DIV dir="<%=dir_obj_var%>" onmouseover="this.style.background='#9CCDF6'" onmouseout="this.style.background='#d3d3d3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#d3d3d3; padding:3px; cursor:hand"
	ONCLICK="parent.location.href='feedbacks.asp?prodId=<%=prodId%>&sort=<%=sort%>&cod='">
    <span id=word19 name=word19><!--הכל--><%=arrTitles(19)%></span>
    </DIV>
</div>
</DIV>
<%set con=nothing%>
</html>
