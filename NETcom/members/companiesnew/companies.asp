<!--  רשימת חברות הכוללת אפשרות חיפוש, מיון ודיפדוף -->
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%search_company = trim(Request.Form("search_company"))
	 if trim(Request.QueryString("search_company")) <> "" then
		search_company = trim(Request.QueryString("search_company"))
	 end if	
	 
	search_city_Name = trim(Request.Form("search_city_Name"))
	if trim(Request.QueryString("search_city_Name")) <> "" then
	search_city_Name = trim(Request.QueryString("search_city_Name"))
	end if		 				
	 
	search_phone = trim(Request.Form("search_phone"))
	if trim(Request.QueryString("search_phone")) <> "" then
	search_phone = trim(Request.QueryString("search_phone"))
	end if	 

	 OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))
	 lang_id = trim(Request.Cookies("bizpegasus")("LANGID"))
	 RowsInList = Request.Cookies("bizpegasus")("RowsInList")
	 
  	 If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 1 
	 End If
	 If lang_id = 2 Then
		dir_var = "rtl"  :  align_var = "left"  : dir_obj_var = "ltr"
		arr_status = Array("","new","active","close","appeal")
	 Else
		dir_var = "ltr"  :  align_var = "right"  :  dir_obj_var = "rtl"
		arr_status = Array("","עתידי","פעיל","סגור","פונה")
	 End If		

	 status = 0
	 if trim(Request.QueryString("status"))<>"" then
		status = cInt(Request.QueryString("status"))		
	 end if 
	 
	 if trim(Request.QueryString("page"))<>"" then
		page=Request.QueryString("page")
	 else
		Page=1
	 end if  
	 
	 if trim(Request.QueryString("numOfRow"))<>"" then
		numOfRow=Request.QueryString("numOfRow")
	 else
		numOfRow = 1
	 end if  
	 
	 sort = 0
	 If isNumeric(Request.QueryString("sort")) And trim(Request.QueryString("sort")) <> "" Then
		sort = cInt(Request.QueryString("sort"))
	 End If 

	urlSort = "companies.asp?search_company="&Server.URLEncode(search_company) & _
	"&amp;search_city_Name=" & Server.URLEncode(search_city_Name) & _
	"&amp;search_phone=" & Server.URLEncode(search_phone) & "&status=" & status

	dim sortby(8)	
	sortby(0) = "company_name"
	sortby(1) = "company_name"
	sortby(2) = "company_name DESC"
	sortby(3) = "city_Name"
	sortby(4) = "city_Name DESC"
	sortby(5) = "CONTACT_NAME"
	sortby(6) = "CONTACT_NAME DESC"
	sortby(7) = "type_name"
	sortby(8) = "type_name DESC"

   If trim(RowsInList) <> "" AND IsNumeric(RowsInList) = true Then
	    PageSize = RowsInList
   Else	
     	PageSize = 10
   End If	    
	
	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 2 Order By word_id"				
	set rstitle = con.getRecordSet(sqlstr)		
	If not rstitle.eof Then
	arrTitlesD = rstitle.getRows()		
	redim arrTitles(Ubound(arrTitlesD,2)+1)
	For i=0 To Ubound(arrTitlesD,2)		
		arrTitles(arrTitlesD(0,i)) = arrTitlesD(1,i)		
	Next
	End If
	set rstitle = Nothing	%>
<html>
<head>
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript">
<!--	
	var oPopup_Status = window.createPopup();
	function StatusDropDown(obj)
	{
	    oPopup_Status.document.body.innerHTML = Status_Popup.innerHTML; 
	    oPopup_Status.show(0-60+obj.offsetWidth, 17, 60, 112, obj);    
	}
	
	function doSearch()
	{
		document.location.href = 'companies.asp?status=<%=status%>' +
		'&search_city_Name=' + document.getElementById("search_city_Name").value + 
		'&search_company=' + document.getElementById("search_company").value + 
		'&search_phone=' + document.getElementById("search_phone").value;			
		return true;
	}	
	
	function entsubsearchg(e) { 
		if( typeof( e ) == "undefined" && typeof( window.event ) != "undefined" )
			e = window.event;

		if (e && e.keyCode == 13)
			return doSearch();		
		else
			return true; }			

//-->
</script>
</head>
<body scroll=no onload="resizeTo(document.body.scrollWidth,document.body.scrollHeight)" style="margin:0px;SCROLLBAR-FACE-COLOR: #E6E6E6;SCROLLBAR-HIGHLIGHT-COLOR: #F7F7F7;SCROLLBAR-SHADOW-COLOR: #848484;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #808080;SCROLLBAR-TRACK-COLOR: #E6E6E6;SCROLLBAR-DARKSHADOW-COLOR: #ffffff">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" valign="top" >
    <table width="100%" cellspacing="1" cellpadding="1" bgcolor="#FFFFFF" dir="<%=dir_var%>">  
	<tr> 	      
  			<td colspan="2" class="title_sort" align="right" dir="ltr"><input type="button" 
			onclick="javascript:document.location.href='companies.asp';" id="btnShowAll" value="הראה הכל" style="width: 70px" 
			class="button" NAME="btnShowAll">&nbsp;&nbsp;<input type="button" onclick="return doSearch();" id="btnSearch" 
			value="חפש" style="width: 50px" class="button" NAME="btnSearch">	  			
  			</td>	      	    
  			<td width="100" nowrap align="center" dir="ltr" class="title_sort"><input type="text" class="search" dir="rtl" style="width:98px;" value="<%=vFix(search_phone)%>" name="search_phone" ID="search_phone" onkeypress="return entsubsearchg(event);"></td>
			<td width="120" nowrap align="center" dir="ltr" class="title_sort"><input type="text" class="search" dir="rtl" style="width:118px;" value="<%=vFix(search_city_Name)%>" name="search_city_Name" ID="search_city_Name" onkeypress="return entsubsearchg(event);"></td>
			<td width="100%" align="right" dir="ltr" class="title_sort"><input type="text" class="search" dir="rtl" style="width:200px;" value="<%=vFix(search_company)%>" name="search_company" ID="search_company" onkeypress="return entsubsearchg(event);" ></td>	     	              
			<td width="44" nowrap class="title_sort">&nbsp;</td>
	</tr>		         
    <tr> 	      
	    <td width="190" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;<!--אימייל--><%=arrTitles(4)%>&nbsp;</td>
	    <td width="90"  nowrap class="title_sort" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<!--פקס--><%=arrTitles(5)%>&nbsp;</td>
  	    <td width="100"  nowrap class="title_sort" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<!--טלפון--><%=arrTitles(6)%>&nbsp;</td>	    
  	    <td width="120"  nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort)="3" OR trim(sort)="4" then%>_act<%end if%>"><%if trim(sort)="3" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" name=word18 title="<%=arrTitles(18)%>"><%elseif trim(sort)="4" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" name=word19 title="<%=arrTitles(19)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort=3" name=word19 title="<%=arrTitles(19)%>"><%end if%><span id="word7" name=word7><!--עיר--><%=arrTitles(7)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort)="3" then%>bot<%elseif trim(sort)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
        <td width="100%" align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort)="1" OR trim(sort)="2" then%>_act<%end if%>"><%if trim(sort)="1" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" name=word18 title="<%=arrTitles(18)%>"><%elseif trim(sort)="2" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" name=word19 title="<%=arrTitles(19)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort=1" name=word19 title="<%=arrTitles(19)%>"><%end if%><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="1" then%>bot<%elseif trim(sort)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	     	              
        <td id=td_status name=td_status width="44" align="<%=align_var%>" class="title_sort" nowrap dir="<%=dir_obj_var%>">&nbsp;<!--סטטוס--><%=arrTitles(8)%>&nbsp;<IMG id=choose3 name=word20 style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 title="<%=arrTitles(20)%>" align=absmiddle onmousedown="StatusDropDown(td_status)"></td>
	</tr>   
<%sqlSelect = "Exec dbo.get_companies @Page=" & Page & ", @RecsPerPage=" & PageSize & ", @OrgID='" & OrgID & _
	"', @company_Name='" & sFix(search_company) & "', @city_name='" & sFix(search_city_Name) & _
	"', @phone='" & sFix(search_phone) & "', @status='" & sFix(status) & "', @sort='" & sort & "'"
     'Response.Write (sqlSelect) 
     'Response.End
     set companyList=con.GetRecordSet(sqlSelect) 
     If Not companyList.EOF then		
		cc = 0
		recCount = companyList("CountRecords")	
		do while not companyList.EOF
			If cc Mod 2 = 0 Then
				tr_bgcolor = "#E6E6E6"
			Else
				tr_bgcolor = "#C9C9C9"
			End If			
			companyID=companyList(1)	  			
			companyName=rtrim(ltrim(companyList(2)))    
			email=trim(companyList(5))
			cityName=trim(companyList(6))
			status_company=trim(companyList(7))
			phone=trim(companyList(3))
			fax=trim(companyList(4)) %>
		<tr bgcolor="<%=tr_bgcolor%>" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='<%=tr_bgcolor%>';">
	      <td valign="top" dir="<%=dir_obj_var%>" align="<%=align_var%>">&nbsp;<%If Len(email) > 0 Then%><a href="mailto:<%=email%>" class="file_link" style="font-size:11px"><%=email%></a><%End If%>&nbsp;</td>
	      <td valign="top" dir="ltr" align="<%=align_var%>"><a class="link_categ" href="company.asp?companyId=<%=companyId%>" target=_parent>&nbsp;<%=fax%>&nbsp;</a></td>
	      <td valign="top" dir="ltr" align="<%=align_var%>"><a class="link_categ" href="company.asp?companyId=<%=companyId%>" target=_parent>&nbsp;<%=phone%>&nbsp;</a></td>	     
	      <td valign="top" dir="<%=dir_obj_var%>" align="<%=align_var%>"><a class="link_categ" href="company.asp?companyId=<%=companyId%>" target=_parent>&nbsp;<%=cityName%>&nbsp;</a></td>
          <td valign="top" dir="<%=dir_obj_var%>" align="<%=align_var%>" dir=rtl><a class="link_categ" href="company.asp?companyId=<%=companyId%>" target=_parent style="line-height:120%;padding-top:3px;padding-bottom:3px">&nbsp;<%=companyName%>&nbsp;</a></td>	                  
          <td align="center"><a class=status_num<%=status_company%> href="company.asp?companyId=<%=companyId%>" target=_parent><%=arr_status(status_company)%></a></td>
      </tr> 
<%	companyList.movenext
	cc = cc+1
	loop
	  
	NumberOfPages = Fix((recCount / PageSize)+0.99)
	if NumberOfPages > 1 then
	urlSort = urlSort & "&sort=" & sort	  
%>
	  <tr>
		<td width="100%" align=middle colspan="8" dir=ltr  bgcolor="#e6e6e6">
			<table border="0" cellspacing="0" cellpadding="2">               
	      <%If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
				Else num = NumberOfPages : numOfRows = 1    	                      
				End If      %>	         
	         <tr>
	         <%if numOfRow <> 1 then%> 
			 <td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A id=word21 name=word21 class=pageCounter title="<%=arrTitles(21)%>" href="<%=urlSort%>&page=<%=10*(numOfRow-1)-9%>&amp;numOfRow=<%=numOfRow-1%>" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	                  if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		                 <td align="middle" bgcolor="#e6e6e6" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter href="<%=urlSort%>&page=<%=i+10*(numOfRow-1)%>&amp;numOfRow=<%=numOfRow%>" ><%=i+10*(numOfRow-1)%></a></td>
	                  <%end if
	                  end if
	               next%>	            
					<td bgcolor="#e6e6e6" align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRow) then%>  
					<td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A id=word22 name=word22 class=pageCounter title="<%=arrTitles(22)%>" href="<%=urlSort%>&page=<%=10*(numOfRow) + 1%>&amp;numOfRow=<%=numOfRow+1%>" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>			
	<%companyList.close 
	set companyList=Nothing
	End if
	%>
	<tr>
	   <td colspan="8" height=18 bgcolor="#e6e6e6" align=center class="card" dir="ltr" style="color:#6F6DA6;font-weight:600"><span id=word9 name=word9><!--נמצאו--><%=arrTitles(9)%></span> <%=recCount%> &nbsp;<%=trim(Request.Cookies("bizpegasus")("CompaniesMulti"))%></td>
	</tr>
	<% 
	Else %>
	   <tr>
	   <td colspan="8" align=center class="title_sort1" dir="<%=dir_var%>"><!--לא נמצאו--><%=arrTitles(10)%>&nbsp;<%=arrTitles(24)%></td>
	   </tr>
<% End If%>
</table></td>
</tr></table>
<DIV ID="Status_Popup" STYLE="display:none;">
<div dir="<%=dir_obj_var%>" style="position:absolute; top:0; left:0; width:60; height:112; border-top:1px solid black; background-color:#d3d3d3;">
<%For i=1 To Ubound(arr_status)%>
	<DIV onmouseover="this.style.background='#6F6DA6'" onmouseout="this.style.background='#d3d3d3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#d3d3d3; border:1px solid black; padding:3px; cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='companies.asp?search_company=<%=Server.URLEncode(search_company)%>&sort=<%=sort%>&status=<%=i%>'">
    <%=arr_status(i)%>
    </DIV>
<%Next%>       
    <DIV onmouseover="this.style.background='#6F6DA6'" onmouseout="this.style.background='#d3d3d3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#d3d3d3; border:1px solid black; padding:3px; cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='companies.asp?search_company=<%=Server.URLEncode(search_company)%>&sort=<%=sort%>'">
    <span id=word15 name=word15><!--הכל--><%=arrTitles(15)%></span>
    </DIV>
</div>
</DIV> 
</body>
<%set con=Nothing%>
</html>