<%@ Page Language="vb" AutoEventWireup="false" Codebehind="AddWorkerMVC.aspx.vb" Inherits="bizpower_pegasus2018.AddWorkerMVC"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>AddWorkerMVC</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
		<script>
		function check_all_bars(objChk,parentID)
	{
		input_arr = document.getElementsByTagName("input");	
		
		for(i=0;i<input_arr.length;i++)	{
		
			
			if(input_arr[i].type == "checkbox")
			{
				currparentId = "";
				objValue = new String(input_arr[i].id);			
				value_arr = objValue.split("!");
				currparentId =  value_arr[1];
				if(currparentId == parentID)
				{
					//input_arr(i).disabled = objChk.checked;
					input_arr[i].checked = objChk.checked;
				}	
			}	
		}
		return true;
	}
	
function check_Parent(objChk,barID,parentID)
{
//alert (objChk.checked +":"+ barID+";"+parentID)
 // alert(document.getElementById("is_visible"+parentID).checked)
if (objChk.checked==true)
{
 if (document.getElementById("is_visible"+parentID).checked==false)
 {
 document.getElementById("is_visible"+parentID).checked=true
 }
}
return true;
}
    
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<table border="0" dir="rtl" align="right" cellspacing="1" cellpadding="1" bgcolor="#E6E6E6"
				width="100%">
				<tr>
							<td height="30" align="center" class="page_title"><span style="COLOR: #6F6DA6;font-size:14pt">הרשאות ביצוע פעולות לעובד  <br><b><%=UserName%></b></span></td>
		</tr>
				
			</table>
			
			<br clear="all">
			<table dir="rtl" align="right" cellspacing="1"	cellpadding="1" bgcolor="#ffffff" width="100%">
			<tr><td valign=top>
			<asp:DataList ID=rptSubSites Runat=server RepeatColumns =2 RepeatDirection="Horizontal" CellSpacing="1"	RepeatLayout="Table" CellPadding="1" BackColor="#f5f5f5" ItemStyle-CssClass="td_admin_4" HorizontalAlign=Right ItemStyle-VerticalAlign=top>
			<ItemTemplate>
			<table border =0 cellpadding=1 cellpadding=1 width =100%>
			<tr><td valign=top class=title_sort><%#Container.Dataitem("SubSitesName")%></td></tR>
			<%'#IIF (Container.DataItem("SubSitesId")=2,"<tr><td>drdddfdf</td></tr>","")%>
			<asp:PlaceHolder ID="plhdep" Runat="server" Visible="False">
			<tr><td bgcolor="#bbbad6"><table cellpadding=1 cellspacing=1>
			<tr><td class="form_titleNormal" align=right bgcolor="#bbbad6" valign=top>מחלקות: </td><td align=center><select runat="server" id="seldep" class="searchList" name="seldep" style="width: 400px;height:100px;direction:rtl;font-size:8pt"
							multiple="">
						</select></td></tr>
						<tr><td height=15 colspan=2>&nbsp;</td></tr>
						</table></td></tr>
			</asp:PlaceHolder>
			<tr><td valign=top><table cellpadding=0 cellspacing=0>
				<asp:Repeater ID="rptSubSiteScreen" Runat="server" OnItemDataBound="CreaterptScreenAction">
				<ItemTemplate>
				<tr><td style="background-color:#bbbad6"><input type="checkbox" dir="ltr" name="is_visibleL<%#Container.DataItem("ScreenId")%>" <%#IIF(Container.DataItem("is_visible")=1,"checked","")%> ID="is_visible<%#Container.DataItem("ScreenId")%>" onclick="return check_all_bars(this,'<%#Container.DataItem("ScreenId")%>')">
						</td>
						
				<td valign=top width="120" align="right" class="form_titleNormal"  dir="rtl" bgcolor="#bbbad6"><%#Container.dataItem("ScreenName")%>&nbsp;</td>
							<td bgcolor="#E6E6E6" align="right"><table cellpadding=0 cellspacing=0 border=0>
							<TR>
								<asp:Repeater ID="rptScreenAction" Runat="server">
								<ItemTemplate>
							<td><input type="checkbox" dir="ltr" name="is_visible<%#Container.DataItem("ActionId")%>" <%#IIF(Container.DataItem("is_visible")=1,"checked","")%>  ID="is_visible<%#Container.DataItem("ActionId")%>!<%#Container.DataItem("ScreenId")%>" onclick="return check_Parent(this,'<%#Container.DataItem("ActionId")%>','<%#Container.DataItem("ScreenId")%>')">
												</td>
												<td align="right" class="title_show_form"><%#Container.dataItem("ActionName")%></td>
			
				
				</ItemTemplate>
				</asp:Repeater>
				</TR>
							</table>
							</td>
			</tr>
			<tr><td colspan=2 height=1 bgcolor=#ffffff></td></tr>
				</ItemTemplate>
				
								</asp:Repeater>
			</table>
			
			</tD></tr>
			</table>
					</ItemTemplate>
			
			</asp:DataList></td></tr>
			<tr><td align=center><table cellpadding=0 cellspacing=0 align=center>
			<tr>
								<td  align="left">
								<asp:Button runat="server" id="Button1" name="Button1" CssClass="but_menu" style="width:90px"></asp:Button>
							
								<td width="150" nowrap></td>
								</td>
								<td align="right"><input type="button" class="but_menu" style="width:90px" onclick="javascript:self.close(); return false;"
										value="ביטול" id="Button2" name="Button2">
								</td>
							</tr>
			</table></td></tr>
			</table>
		</form>
	</body>
</HTML>
