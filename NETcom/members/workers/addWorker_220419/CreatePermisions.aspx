<%@ Page Language="vb" AutoEventWireup="false" Codebehind="CreatePermisions.aspx.vb" Inherits="bizpower_pegasus2018.CreatePermisions"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
    <title>CreatePermisions</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../IE4.css" rel="STYLESHEET" type="text/css">
  </head>
  <body MS_POSITIONING="GridLayout">

    <form id="Form1" method="post" runat="server">
<table border=0 cellpadding=0 cellspacing=0 align=center  width=100%>

<tr>
								<td height="30" align="center"><span style="COLOR: #6F6DA6;font-size:14pt"><b><%=UserName%></b> ניהול הרשאה למתן הרשאות</span></td>
							</tr>
							<tr><td height=20></td></tr>
	<tr>
															<TD align=center><select runat="server" id="seldep" class="searchList" name="seldep" style="width: 220px;height:60px;direction:rtl;font-size:8pt"
																	multiple="">
																</select></TD>
														</tr>

		<tr><td height=20></td></tr>
							
</table>
<table border=0 cellpadding=0 cellspacing=0  align=right ID="Table13" style="background-color:#ffffff" width=100%>
<tr><td class="title" align="right">הרשאות&nbsp;&nbsp;<span style="font-weight:500">ביצוע פעולות</span></td></tr>

<tr><td valign=top bgcolor="#e6e6e6">

			<table border=0 cellpadding=1 cellspacing=1  align=right ID="Table13" style="background-color:#ffffff" width=100%>
			<tr>
				<td class="title_sort" align=center width=14%>טיים לימיטים</td>
				<td class="title_sort" align =center width=14%>מסך טפסים וסקרים -<br> דוח מעקב רישום ואחוזי סגירה</td>
				<td class="title_sort" align=center width=14%>מסך אופרציה</td>
				<td class="title_sort" align=center width=16% nowrap>מסך משובים</td>
				<td class="title_sort" align=center width=14%>מסך מיקודי טיולים</td>
				<td class="title_sort" align=center width=14%>דש בורד מכירות</td>
				<TD class="title_sort" align =center width=14%>כללי</TD>
    		</tr>

    	<tr>
			<td class="title_sort" align=center width=14%>טיים לימיטים</td>
			<td class="title_sort" align =center width=14%>מסך טפסים וסקרים -<br> דוח מעקב רישום ואחוזי סגירה</td>
			<td class="title_sort" align=center width=14%>מסך אופרציה</td>
			<td class="title_sort" align=center width=16% nowrap>מסך משובים</td>
			<td class="title_sort" align=center width=14%>מסך מיקודי טיולים</td>
			<td class="title_sort" align=center width=14%>דש בורד מכירות</td>
<!--clali-->
	<td valign=top bgcolor="#e6e6e6">
			<table border=0 cellpadding=0 cellspacing=0  align=right ID="Table9">
			<tr>
				<td align="right" class="title_show_form" nowrap dir=rtl valign=top>חברת תעופה&nbsp;&nbsp;</td>
				<td align="right" nowrap><input type="checkbox" dir="ltr" name="EditFlyCompanies"  ID="EditFlyCompanies"></td>
				</tr>

			</table>
    </td>
		<!--clali-->
    		</tr>
	     	</table>
</td></tr>
<!---block 2-->


<tr><td height=5 nowrap ></td></tr>
<tr><td class="title" align="right"><!--הרשאות-->הרשות כניסה למודולים&nbsp;&nbsp;<span style="font-weight:500"></span></td></tr>
<tr><td valign=top bgcolor="#e6e6e6">
<table border=0 cellpadding=0 cellspacing=1  align=right dir=rtl width=100%>
<tr><td width=100%><table border=0 cellpadding=0 cellspacing=1  align=right dir=rtl bgcolor=#ffffff width=100%>

<asp:Repeater ID=rptParent Runat=server>
<ItemTemplate><tr><td  align=right width=100% bgcolor="#e6e6e6"><table cellpadding=0 cellspacing=0  align=right border=0 width=100% ><tr><td><table cellpadding=0 cellspacing=0 >
<tr>
<td class="form_titleNormal" align="right" bgcolor=#e4e4e4><input type=checkbox dir="ltr" name="is_visible<%#Container.dataItem("Permission_Id")%>" ID="is_visible<%#Container.dataItem("Permission_Id")%>" ></td>
<td align="right" class="form_titleNormal" bgcolor=#bbbad6 width=80><%#Container.dataItem("Permission_Name")%>&nbsp;</td>

<asp:Repeater ID=rptBar Runat=server>
<ItemTemplate>
	<td align="right"><input type=checkbox dir="ltr" name="is_visible<%#Container.dataItem("Permission_Id")%>" id="xx" ></td>
	<td align="right" class="title_show_form"  nowrap dir="rtl"><%#Container.dataItem("Permission_Name")%>&nbsp;</td>

</ItemTemplate>
</asp:Repeater>
</tr>
</table>
</td></tr></table></td></tr>
</ItemTemplate>
</asp:Repeater>
</table></td></tr>
</table>
</td></tr>

</table>
    </form>

  </body>
</html>
