<%@ Page Language="vb" AutoEventWireup="false" Codebehind="contacts_privateQ16504.aspx.vb" Inherits="bizpower_pegasus2018.contacts_privateQ16504"%>
<%@ Register TagPrefix="DS" TagName="metaInc" Src="../../title_meta_inc.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<title>טופס מתעניין בטיול</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<DS:metaInc id="metaInc" runat="server"></DS:metaInc>
		<meta charset="utf-8">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</head>
	<BODY style="margin:0px;SCROLLBAR-FACE-COLOR: #E6E6E6;SCROLLBAR-HIGHLIGHT-COLOR: #F7F7F7;SCROLLBAR-SHADOW-COLOR: #848484;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #808080;SCROLLBAR-TRACK-COLOR: #E6E6E6;SCROLLBAR-DARKSHADOW-COLOR: #ffffff'"
		onload='parent.document.getElementById("contact_name").focus()'>
		<form id="Form1" method="post" runat="server">
			<table width="100%" align="center" cellpadding="0" cellspacing="0" border="0" bgcolor="#E6E6E6">
				<tr>
					<td align="left" width="100%" valign=top dir="<%=dir_var%>" >
						<table width="100%" cellspacing="1" cellpadding="1" bgcolor="#FFFFFF">
							<tr>
								<td width="75"  nowrap class="title_sort" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<span id="word6" name="word6"><!--טלפון ישיר--><%=arrTitles.rows(6)(1)%></span></td>
								<td width="120"  nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort)="5" OR trim(sort)="6" then%>_act<%end if%>"><%if trim(sort)="5" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" name=word17 title="<%=arrTitles.rows(17)(1)%>"><%elseif trim(sort)="6" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" name=word18 title="<%=arrTitles.rows(18)(1)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort=5" name=word18 title="<%=arrTitles.rows(18)(1)%>"><%end if%><%=HttpUtility.UrlDecode(Request.Cookies("bizpegasus")("ContactsOne"), System.Text.Encoding.GetEncoding("windows-1255"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="5" then%>bot<%elseif trim(sort)="6" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
							</tr>
							<asp:Repeater ID="rptContacts" Runat="server">
								<ItemTemplate>
									<tr>
										<td class="card" align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" href="TripForm16504.aspx?quest_id=<%=quest_id%>&companyId=<%=companyId%>&contactId=<%#container.dataitem("contact_Id")%>&isSelectedContact=1" target=_parent style="font-size:11px">&nbsp;<%#container.dataitem("Phone")%>&nbsp;</a></td>
										<td class="card" align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" href="TripForm16504.aspx?quest_id=<%=quest_id%>&companyId=<%=companyId%>&contactId=<%#container.dataitem("contact_Id")%>&isSelectedContact=1" target=_parent>&nbsp;<%#container.dataitem("contact_Name")%>&nbsp;</a></td>
									</tr>
								</ItemTemplate>
							</asp:Repeater>
						</table>
					</td>
				</tr>
			</table>
		</form>
	</BODY>
</html>
