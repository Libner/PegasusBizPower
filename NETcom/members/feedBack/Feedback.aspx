<%@ Page Language="vb" AutoEventWireup="false" Codebehind="Feedback.aspx.vb" Inherits="bizpower_pegasus2018.Feedback" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Feedback</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
	</HEAD>
	<body style="MARGIN:0px">
		<form id="Form1" method="post" runat="server">
			<table border="0" cellpadding="0" cellspacing="0" width="100%" align="left">
				<TBODY>
					<tr>
						<td align="center" style="FONT-SIZE:14px;FONT-WEIGHT:bold;COLOR:#000000;BACKGROUND-COLOR:#ffd011">מסך 
							משוב של <%=CONTACTNAME%>
						</td>
					</tr>
					<tr>
						<td>
							<%if status=2 then%>
							<table border="0" cellpadding="0" cellspacing="0" width="100%" align="left">
								<tr bgcolor="#d8d8d8" style="HEIGHT:25px">
									<td class="title_sort" align="center">ציון המשוב</td>
									<td class="title_sort" align="center">ציון הטיול</td>
									<td class="title_sort" align="center">ציון ממוצע של מדריך למשוב זה</td>
									<td class="title_sort" align="center">מדריך הטיול</td>
									<td class="title_sort" align="center">קוד הטיול</td>
									<td class="title_sort" align="center">שם הטיול</td>
								</tr>
								<tr style="HEIGHT:25px">
									<td align="center"><%=TourGrade%>%</td>
									<td align="center"><%=func.TourGrade(DepartureId)%></td>
									<td align="center"><%=GuideGrade%>%</td>
									<th align="center" class="title_show_form">
										<%=GuideName%>
									</th>
									<th align="center" class="title_show_form">
										<%=DepartureCode%>
									</th>
									<th align="center" class="title_show_form">
										<%=TourName%>
									</th>
								</tr>
							</table>
							<%elseif status=1 then%>
							<table border="0" cellpadding="0" cellspacing="0" width="100%" align="left">
								<tr bgcolor="#d8d8d8" style="HEIGHT:25px">
									<td align="right" class="title_sort" width="10%">מספר הפעמים שנשלחה ההודעה ללקוח 
										באמצעי כלשהו</td>
									<td align="right" class="title_sort" width="10%" dir="rtl">האם פתח הלינק דרך טאבלט 
										או דרך מחשב ביתי?</td>
									<td align="right" class="title_sort" width="10%" dir="rtl">באיזה חלק של המשוב הפסיק 
										הלקוח?</td>
									<td class="title_sort" align="center">מדריך הטיול</td>
									<td class="title_sort" align="center">קוד הטיול</td>
									<td class="title_sort" align="center">שם מלא של המטייל</td>
								</tr>
								<tr style="HEIGHT:25px">
									<td align="center"><%=SendCount%></td>
									<td align="center"><%=OpenDevice%></td>
									<td align="center"><%=LastFAQ%>
										שאלה</td>
									<th align="center" class="title_show_form">
										<%=GuideName%>
									</th>
									<th align="center" class="title_show_form">
										<%=DepartureCode%>
									</th>
									<th align="center" class="title_show_form">
										<%=TourName%>
									</th>
								</tr>
							</table>
							<br>
							<br clear="all">
							<br>
							<br clear="all">
							<%end if%>
							<br>
							<br clear="all">
							<br>
							<br clear="all">
							<asp:Repeater ID="rptCat" Runat="server">
								<HeaderTemplate>
									<table cellpadding="3" cellspacing="1" width="100%" align="center">
								</HeaderTemplate>
								<ItemTemplate>
									<tr style="background-color: rgb(201, 201, 201); height: 30px; " onmouseover="this.style.backgroundColor='#B1B0CF';"
										onmouseout="this.style.backgroundColor='#C9C9C9';">
										<td align="right" colspan="4"><b><%#Container.dataItem("CategoryName")%></b>&nbsp;</td>
									</tr>
									<asp:Repeater ID="rptFaq" Runat="server">
										<ItemTemplate>
											<tr>
												<td align="right" dir="rtl" width=100%>
													<asp:Label runat="server" ID="lblValue" Text="" ForeColor="black"></asp:Label></td>
												<TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
												<td align="right" dir="rtl" nowrap valign=top>
											
													<asp:Label runat="server" ID="lblTitle" Text="" ForeColor="black"></asp:Label><%'#Container.dataItem("FAQ_Title")%></td>
												<td align="right" nowrap valign=top>
													<asp:Label runat="server" ID="lblFAQID" Text="" ForeColor="black"></asp:Label></td>
											</tr>
										</ItemTemplate>
										<SeparatorTemplate>
											<tr style="background-color: rgb(201, 201, 201);height:1px;">
												<td align="right" colspan="4" style="height:1px;" width="100%">
											</tr>
										</SeparatorTemplate>
									</asp:Repeater>
								</ItemTemplate>
								<FooterTemplate>
			</table>
			</FooterTemplate> </asp:Repeater>
			<%if false then%>
			<asp:PlaceHolder Runat="server" ID="plhStatus2" Visible="false">
				<TABLE cellSpacing="1" cellPadding="1" width="50%" align="center" border="0">
					<TR style="HEIGHT: 25px" bgColor="#d8d8d8">
						<TH class="title_show_form" width="10%" align="right">
							ציון כללי של המשוב&nbsp;</TH>
						<TH class="title_show_form" width="10%" align="right">
							קטגוריית הציון השני הכי גרוע שהלקוח נתן&nbsp;</TH>
						<TH class="title_show_form" width="10%" align="right">
							קטגוריית הציון הכי גרוע שהלקוח נתן&nbsp;</TH>
						<TH class="title_show_form" width="10%" align="right">
							שביעות רצון כללית מהטיול&nbsp;14</TH>
						<TH class="title_show_form" width="50%" align="right">
							בהנחה ויתאפשר, באיזה מידה סביר כי תשוב לטייל עם פגסוס&nbsp;13</TH></TR>
					<TR onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#C9C9C9';"
						style="HEIGHT: 30px; BACKGROUND-COLOR: rgb(201,201,201)">
						<TD align="center">1</TD>
						<TD align="center"><%=func.Feedback_GetNextMinGrade(DepartureId,ContactId)%></TD>
						<TD align="center"><%=FaqGradeMinValue%></TD>
						<TD align="center"><%=FAQ14%>
						<TD align="right"><%=FAQ13%></TD>
					</TR>
				</TABLE>
			</asp:PlaceHolder><%end if%>
		</form>
		</TD></TR></TBODY></TABLE>
	</body>
</HTML>
