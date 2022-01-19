<%@ Page Language="vb" AutoEventWireup="false" Codebehind="CreatePermisions_withoutFieldsActions.aspx.vb" Inherits="bizpower_pegasus2018.CreatePermisions"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<title>CreatePermisions</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
		<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
		<script type="text/javascript" src="//code.jquery.com/jquery-1.12.4.js"></script>
		<script>
    $(document).ready(function()
{ 
			$('#Button1').click(
		function()
		{

				if (!$("#seldep option:selected").length)
				{
				alert ("אנא בחר מחלקה")
				return false;
				}
		}
		);
}



);
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
	

    
		</script>
	</head>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<table border="0" cellpadding="0" cellspacing="0" align="center" width="100%">
				<tr>
					<td height="30" align="center" dir=rtl><span style="COLOR: #6F6DA6;font-size:14pt">ניהול הרשאה למתן הרשאות לעובד  <br><b><%=UserName%></b></span></td>
				</tr>
				<tr>
					<td height="20"></td>
				</tr>
				<tr>
					<TD align="center"><select runat="server" id="seldep" class="searchList" name="seldep" style="width: 220px;height:110px;direction:rtl;font-size:8pt"
							multiple="">
						</select></TD>
				</tr>
				<tr>
					<td height="20"></td>
				</tr>
			</table>
			<table border="0" cellpadding="0" cellspacing="0" align="right" ID="Table13" style="background-color:#ffffff;display:block"
				width="100%">
				<tr>
					<td class="title" align="right">הרשאות&nbsp;&nbsp;<span style="font-weight:500">ביצוע 
							פעולות</span></td>
				</tr>
				<tr>
					<td valign="top" bgcolor="#e6e6e6">
						<table border="0" cellpadding="1" cellspacing="1" align="right" ID="Table13" style="background-color:#ffffff"
							width="100%">
							<tr>
								<td class="title_sort" align="center" width="14%">טיים לימיטים</td>
								<td class="title_sort" align="center" width="14%">מסך טפסים וסקרים -<br>
									דוח מעקב רישום ואחוזי סגירה</td>
								<td class="title_sort" align="center" width="14%">מסך אופרציה</td>
								<td class="title_sort" align="center" width="16%" nowrap>מסך משובים</td>
								<td class="title_sort" align="center" width="14%">מסך מיקודי טיולים</td>
								<td class="title_sort" align="center" width="14%">דש בורד מכירות</td>
								<TD class="title_sort" align="center" width="14%">כללי</TD>
							</tr>
							<tr>
								<td valign="top" bgcolor="#e6e6e6">
									<table border="0" cellpadding="0" cellspacing="0" align="right" ID="Table9">
										<asp:Repeater ID="rpt90007" Runat="server">
											<ItemTemplate>
												<tr>
													<td align="right" class="title_show_form" nowrap dir="rtl" valign="top"><%#Container.DataItem("Permission_Name")%>&nbsp;&nbsp;</td>
													<td align="right" nowrap><input type="checkbox" dir="ltr" name="is_visible<%#Container.DataItem("Bar_Id")%>"  ID="is_visible<%#Container.DataItem("Bar_Id")%>" <%#IIF(Container.DataItem("is_visible")=1,"checked","")%>></td>
												</tr>
											</ItemTemplate>
										</asp:Repeater>
									</table>
								</td>
								<td valign="top" bgcolor="#e6e6e6">
									<table border="0" cellpadding="0" cellspacing="0" align="right" ID="Table9">
										<asp:Repeater ID="rpt90006" Runat="server">
											<ItemTemplate>
												<tr>
													<td align="right" class="title_show_form" nowrap dir="rtl" valign="top"><%#Container.DataItem("Permission_Name")%>&nbsp;&nbsp;</td>
													<td align="right" nowrap><input type="checkbox" dir="ltr" name="is_visible<%#Container.DataItem("Bar_Id")%>"  ID="is_visible<%#Container.DataItem("Bar_Id")%>" <%#IIF(Container.DataItem("is_visible")=1,"checked","")%>></td>
												</tr>
											</ItemTemplate>
										</asp:Repeater>
									</table>
								</td>
								<td valign="top" bgcolor="#e6e6e6">
									<table border="0" cellpadding="0" cellspacing="0" align="right" ID="Table9">
										<asp:Repeater ID="rpt90005" Runat="server">
											<ItemTemplate>
												<tr>
													<td align="right" class="title_show_form" nowrap dir="rtl" valign="top"><%#Container.DataItem("Permission_Name")%>&nbsp;&nbsp;</td>
													<td align="right" nowrap><input type="checkbox" dir="ltr" name="is_visible<%#Container.DataItem("Bar_Id")%>"  ID="is_visible<%#Container.DataItem("Bar_Id")%>" <%#IIF(Container.DataItem("is_visible")=1,"checked","")%>></td>
												</tr>
											</ItemTemplate>
										</asp:Repeater>
									</table>
								</td>
								<td valign="top" bgcolor="#e6e6e6">
									<table border="0" cellpadding="0" cellspacing="0" align="right" ID="Table9">
										<asp:Repeater ID="rpt90004" Runat="server">
											<ItemTemplate>
												<tr>
													<td align="right" class="title_show_form" nowrap dir="rtl" valign="top"><%#Container.DataItem("Permission_Name")%>&nbsp;&nbsp;</td>
													<td align="right" nowrap><input type="checkbox" dir="ltr" name="is_visible<%#Container.DataItem("Bar_Id")%>"  ID="is_visible<%#Container.DataItem("Bar_Id")%>" <%#IIF(Container.DataItem("is_visible")=1,"checked","")%>></td>
												</tr>
											</ItemTemplate>
										</asp:Repeater>
									</table>
								</td>
								<td valign="top" bgcolor="#e6e6e6">
									<table border="0" cellpadding="0" cellspacing="0" align="right" ID="Table9">
										<asp:Repeater ID="rpt90003" Runat="server">
											<ItemTemplate>
												<tr>
													<td align="right" class="title_show_form" nowrap dir="rtl" valign="top"><%#Container.DataItem("Permission_Name")%>&nbsp;&nbsp;</td>
													<td align="right" nowrap><input type="checkbox" dir="ltr" name="is_visible<%#Container.DataItem("Bar_Id")%>"  ID="is_visible<%#Container.DataItem("Bar_Id")%>" <%#IIF(Container.DataItem("is_visible")=1,"checked","")%>></td>
												</tr>
											</ItemTemplate>
										</asp:Repeater>
									</table>
								</td>
								<td valign="top" bgcolor="#e6e6e6">
									<table border="0" cellpadding="0" cellspacing="0" align="right" ID="Table9">
										<asp:Repeater ID="rpt90002" Runat="server">
											<ItemTemplate>
												<tr>
													<td align="right" class="title_show_form" nowrap dir="rtl" valign="top"><%#Container.DataItem("Permission_Name")%>&nbsp;&nbsp;</td>
													<td align="right" nowrap><input type="checkbox" dir="ltr" name="is_visible<%#Container.DataItem("Bar_Id")%>"  ID="is_visible<%#Container.DataItem("Bar_Id")%>" <%#IIF(Container.DataItem("is_visible")=1,"checked","")%>></td>
												</tr>
											</ItemTemplate>
										</asp:Repeater>
										<tr><td height=30 colspan=2>&nbsp;</td></tr>
																<tr><td width=100% colspan=2 align=right><table cellspacing="1" cellpadding="1" width=100%>
																<tr><td width="100%" align="center" class="title_sort" colspan=2 height=25>מטיילים</td></tr>
															<asp:Repeater ID="rpt90097" Runat="server">
											<ItemTemplate>
												<tr>
													<td align="right" class="title_show_form" nowrap dir="rtl" valign="top"><%#Container.DataItem("Permission_Name")%>&nbsp;&nbsp;</td>
													<td align="right" nowrap><input type="checkbox" dir="ltr" name="is_visible<%#Container.DataItem("Bar_Id")%>"  ID="is_visible<%#Container.DataItem("Bar_Id")%>" <%#IIF(Container.DataItem("is_visible")=1,"checked","")%>></td>
												</tr>
											</ItemTemplate>
										</asp:Repeater>
																</table></td></tr>

									</table>
								</td>
								<!--clali-->
								<td valign="top" bgcolor="#e6e6e6">
									<table border="0" cellpadding="0" cellspacing="0" align="right" ID="Table9">
										<asp:Repeater ID="rpt90001" Runat="server">
											<ItemTemplate>
												<tr>
													<td align="right" class="title_show_form" nowrap dir="rtl" valign="top"><%#Container.DataItem("Permission_Name")%>&nbsp;&nbsp;</td>
													<td align="right" nowrap><input type="checkbox" dir="ltr" name="is_visible<%#Container.DataItem("Bar_Id")%>"  ID="is_visible<%#Container.DataItem("Bar_Id")%>" <%#IIF(Container.DataItem("is_visible")=1,"checked","")%>></td>
												</tr>
											</ItemTemplate>
										</asp:Repeater>
									</table>
								</td>
								<!--clali-->
							</tr>
						</table>
					</td>
				</tr>
				<!---block 2-->
				<tr>
					<td height="5" nowrap></td>
				</tr>
				<tr>
					<td class="title" align="right" dir=rtl><!--הרשאות--> הרשאות כניסה למודולים&nbsp;&nbsp;<span style="font-weight:500">(סמן בקוביה הימנית את המודול בו העובד רשאי לטפל ובהמשך לאיזה חלקים מתוכו יקבל הרשאה) </span></td>
				</tr>
				<tr>
					<td valign="top" bgcolor="#e6e6e6">
						<table border="0" cellpadding="0" cellspacing="1" align="right" dir="rtl" width="100%">
							<tr>
								<td width="100%"><table border="0" cellpadding="0" cellspacing="1" align="right" dir="rtl" bgcolor="#ffffff"
										width="100%">
										<asp:Repeater ID="rptParent" Runat="server">
											<ItemTemplate>
												<tr>
													<td align="right" width="100%" bgcolor="#e6e6e6"><table cellpadding="0" cellspacing="0" align="right" border="0" width="100%">
															<tr>
																<td><table cellpadding="0" cellspacing="0">
																		<tr>
																			<td class="form_titleNormal" align="right" bgcolor="#e4e4e4"><input type=checkbox dir="ltr" name="is_visible<%#Container.DataItem("Bar_Id")%>" ID="is_visible<%#Container.DataItem("Bar_Id")%>" <%#IIF(Container.DataItem("is_visible")=1,"checked","")%> onclick="return check_all_bars(this,'<%#Container.DataItem("Bar_Id")%>')"></td>
																			<td align="right" class="form_titleNormal" bgcolor="#bbbad6" width="80"><%#Container.dataItem("Permission_Name")%>&nbsp;</td>
																			<asp:Repeater ID="rptBar" Runat="server">
																				<ItemTemplate>
																					<td align="right"><input type=checkbox dir="ltr"  name="is_visible<%#Container.DataItem("bar_Id")%>" id="<%#Container.DataItem("bar_Id")%>!<%#Container.DataItem("Parent_Id")%>" <%#IIF(Container.DataItem("is_visible")=1,"checked","")%> onclick="return check_Parent(this,'<%#Container.DataItem("Bar_Id")%>','<%#Container.DataItem("Parent_Id")%>')"></td>
																					<td align="right" class="title_show_form" nowrap dir="rtl"><%#Container.dataItem("Permission_Name")%>&nbsp;</td>
																				</ItemTemplate>
																			</asp:Repeater>
																		</tr>
																	</table>
																</td>
															</tr>
														</table>
													</td>
												</tr>
											</ItemTemplate>
										</asp:Repeater>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td><table border="0" cellpadding="0" cellspacing="0" align="left" ID="Table1" width="100%">
							<tr>
								<td height="5" nowrap colspan="3"></td>
							</tr>
							<tr>
								<td height="1" nowrap colspan="3" bgcolor="#808080"></td>
							</tr>
							<tr>
								<td height="10" nowrap colspan="3"></td>
							</tr>
							<tr>
								<td width="50%" align="right"><input type="button" class="but_menu" style="width:90px" onclick="javascript:self.close(); return false;"
										value="ביטול" id="Button2" name="Button2">
								</td>
								<td width="150" nowrap></td>
								<td width="50%" align="left">
								<asp:Button runat="server" id="Button1" name="Button1" CssClass="but_menu" style="width:90px"></asp:Button>
								</td>
							</tr>
							<tr>
								<td height="5" nowrap colspan="3"></td>
							</tr>
							<tr>
								<td height="1" nowrap colspan="3" bgcolor="#808080"></td>
							</tr>
							<tr>
								<td height="10" nowrap colspan="3"></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>
