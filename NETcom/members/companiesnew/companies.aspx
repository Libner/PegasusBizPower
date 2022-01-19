<%@ OutputCache Duration="1" VaryByParam="none" Location="none"%>
<%@ Page Language="vb" AutoEventWireup="false" Codebehind="companies.aspx.vb" Inherits="bizpower_pegasus.companies" %>
<%@ Register TagPrefix="DS" TagName="TopIn" Src="../../top_in.ascx" %>
<%@ Register TagPrefix="DS" TagName="LogoTop" Src="../../logo_top.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>BizPower</title>
		<meta charset="windows-1255">
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
		<SCRIPT LANGUAGE="javascript">
<!--
	var oPopup_City = window.createPopup();
	function CityDropDown(obj){
	    oPopup_City.document.body.innerHTML = City_Popup.innerHTML; 
	    oPopup_City.document.charset="windows-1255";
	    var popupBodyObj = oPopup_City.document.body;   
	    var dHeight = 142;
	    oPopup_City.show(0, 17, obj.offsetWidth, dHeight, obj);    
	}
	
	var oPopup_Type = window.createPopup();
	function company_typeDropDown(obj){
	    oPopup_Type.document.body.innerHTML = company_type_Popup.innerHTML; 
	    oPopup_Type.document.charset="windows-1255";
	    oPopup_Type.show(0, 17, obj.offsetWidth, 148, obj);    
	}
	
	var oPopup_Status = window.createPopup();
	function StatusDropDown(obj)
	{
	    oPopup_Status.document.body.innerHTML = Status_Popup.innerHTML; 
	    oPopup_Status.show(0-60+obj.offsetWidth, 17, 60, 90, obj);    
	}

function DoCal(elTarget){
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("../../calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=134pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		return false;
		window.document.focus;   
}
//-->
		</SCRIPT>
	</HEAD>
	<body>
		<form id="Form1" method="post" runat="server">
		<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr>
				<td width=100% align="<%=align_var%>">
					<DS:LOGOTOP id="logotop" runat="server"></DS:LOGOTOP>
				</td>
			</tr>
			<tr>
				<td width=100% align="<%=align_var%>">
					<DS:TOPIN id="topin" numOfLink="0" numOfTab="0" runat="server"></DS:TOPIN>
				</td>
			</tr>
			<tr><td class="page_title">&nbsp;</td></tr>
			
			<tr>    
				<td width="100%" valign="top" align="center">
				<table dir="<%=dir_var%>" border="0" cellpadding="0" cellspacing="1" width="100%">
					<tr>
						<td align="left" width="100%" valign=top >
						<table width="100%" cellspacing="1" cellpadding="1" bgcolor=#FFFFFF dir="<%=dir_var%>">       
							<tr> 	      
								<td id=td_group name=td_group width="105" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;<span id=word3 name=word3>קבוצה</span>&nbsp;<IMG id=choose1 name=word20 id="find_stat" style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 title="בחר מרשימה" align=absmiddle onclick="return false" onmousedown="company_typeDropDown(td_group)"></td>
								<td width="150" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;<span id="word4" name=word4>אימייל</span>&nbsp;</td>
								<td width="75"  nowrap class="title_sort" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<span id="word5" name=word5>פקס</span>&nbsp;</td>
  								<td width="75"  nowrap class="title_sort" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<span id="word6" name=word6>טלפון</span>&nbsp;</td>	    
  								<td width="90"  nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort)="3" OR trim(sort)="4" then%>_act<%end if%>"><%if trim(sort)="3" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" name=word18 title="למיון בסדר יורד"><%elseif trim(sort)="4" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" name=word19 title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&sort=3" name=word19 title="למיון בסדר עולה"><%end if%><span id="word7" name=word7>עיר</span>&nbsp;<img src="../../images/arrow_<%if trim(sort)="3" then%>bot<%elseif trim(sort)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
								<td width="100%" align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort)="1" OR trim(sort)="2" then%>_act<%end if%>"><%if trim(sort)="1" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" id=sort_down1 name=word18 title="למיון בסדר יורד"><%elseif trim(sort)="2" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" id=sort_up1 name=word19 title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&sort=1" id=sort_up1 name=word19 title="למיון בסדר עולה"><%end if%><%=HttpContext.Current.Server.UrlDecode(trim(Request.Cookies("bizpegasus")("CompaniesOne")))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="1" then%>bot<%elseif trim(sort)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	     	              
								<td id=td_status name=td_status width="44" align="<%=align_var%>" class="title_sort" nowrap dir="<%=dir_obj_var%>">&nbsp;<span id="word8" name=word8>סטטוס</span>&nbsp;<IMG id=choose3 name=word20 style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 title="בחר מרשימה" align=absmiddle onmousedown="StatusDropDown(td_status)"></td>
							</tr>  
							<asp:PlaceHolder id="CompanyListPH" runat="server"></asp:PlaceHolder>
							<asp:PlaceHolder id="PagingPH" runat="server"></asp:PlaceHolder>
	  
						<%if recCount > 0 then%>
							<tr>
							<td colspan="7" height=18 bgcolor="#e6e6e6" align=center class="card" dir="ltr" style="color:#6F6DA6;font-weight:600"><span id=word9 name=word9>נמצאו</span> <%=recCount%> &nbsp;<%=HttpContext.Current.Server.UrlDecode(trim(Request.Cookies("bizpegasus")("CompaniesMulti")))%></td>
							</tr>
						<%Else %>
							<tr>
							<td colspan="7" align=center class="title_sort1" dir="<%=dir_var%>"><span id="word10" name=word10>לא נמצאו</span> &nbsp;</td>
							</tr>
						<% End If%>
						</table>
					</td>
					<td width=100 nowrap valign=top class="td_menu" style="border: 1px solid #808080" dir="<%=dir_var%>">
						<table cellpadding=1 cellspacing=0 width=100% dir="<%=dir_var%>">
							<tr><td align="<%=align_var%>" colspan=2 class="title_search"><span id=word11 name=word11>חיפוש</span></td></tr>
							<tr><td colspan=2 align="<%=align_var%>" class="right_bar"><%=HttpContext.Current.Server.UrlDecode(trim(Request.Cookies("bizpegasus")("CompaniesOne")))%></td></tr>
							<tr>
							<td align="<%=align_var%>"><input type="image" onclick="Form1.action='companies.aspx?sort=<%=sort%>';Form1.submit();" src="../../images/search.gif" ID="Image1" NAME="Image1"></td>
							<td align="<%=align_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:80;" value="<%=search_company%>" name="search_company"></td></tr>
							<tr><td colspan=2 align="<%=align_var%>" class="right_bar"><%=HttpContext.Current.Server.UrlDecode(trim(Request.Cookies("bizpegasus")("ContactsOne")))%></td></tr>
							<tr><td align="<%=align_var%>"><input type="image" onclick="Form1.action='contacts.asp?sort=<%=sort%>';Form1.submit();" src="../../images/search.gif" ID="Image2" NAME="Image2"></td>
							<td align="<%=align_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:80;" value="<%=search_contact%>" name="search_contact" ID="search_contact"></td></tr>
							<tr><td colspan=2 height=10 nowrap></td></tr>
							<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:95;" href="#" onclick="javascript:window.open('editcompany.asp','NewCompany','top=50,left=100,resizable=0,width=450,height=480');"><span id=word12 name=word12>הוסף</span> <%=HttpContext.Current.Server.UrlDecode(trim(Request.Cookies("bizpegasus")("CompaniesOne")))%></a></td></tr>
							<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:95;" href="#" onclick="javascript:window.open('newcontact.asp','NewContact','top=50,left=40,resizable=0,width=700,height=450');"><span id="word13" name=word13>הוסף</span> <%=HttpContext.Current.Server.UrlDecode(trim(Request.Cookies("bizpegasus")("ContactsOne")))%></a></td></tr>
							<tr><td colspan=2 height=10 nowrap></td></tr>
							<asp:PlaceHolder id="FormsPH" runat="server"></asp:PlaceHolder>  	
						</table>
					</td>
				</tr>  
			</table>
			</td>
			</tr> 
			</table> 
		</form>
<asp:PlaceHolder id="DivPH" runat="server"></asp:PlaceHolder>  
<asp:PlaceHolder id="ScriptPH" runat="server"></asp:PlaceHolder>  
	</body>
</HTML>
