<%@ Page Language="vb" AutoEventWireup="false" Codebehind="info.aspx.vb" Inherits="bizpower_pegasus.info" %>
<%@ OutputCache Duration="30" VaryByParam="None" VaryByCustom="xmlChanged;cookieCode" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
  <head>
    <title>info</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <style type="text/css" media="all">
    TD.title_small { color: #000000; font-family: Arial; font-size: 8pt; }
    TD.value { font-family: Arial; font-size: 8pt; color: #FFFFFF;  font-weight: bold; }
    TD.table_title { font-family: Arial; font-size: 10pt; color: #C6D7E5;  font-weight: bold; }
    </style>
  </head>
  <body style="margin: 0px; overflow: hidden;">
    <form id="Form1" method="post" runat="server" style="margin: 0px;">
    <table border="0" width="766" height="62" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" width="53"><a href="report_orders.aspx" target="_blank" title="דוח הזמנות לנציג"><img src="images/info/but_doh.gif" width="49" height="54" border="0"></a></td>
		<td valign="top" width="170"><table border="0" width="170" height="62" cellspacing="0" cellpadding="0">
		<tr>
			<td width="100%" background="images/info/title_bgr.gif" height="19" align="right" class="table_title">הבונוסים שלי&nbsp;&nbsp; </td>
		</tr>
		<tr>
			<td width="100%" bgcolor="#F4F4F4" height="2"></td>
		</tr>
		<tr>
			<td width="100%" bgcolor="#F4F4F4" height="16" align="right"><table border="0" width="170" cellspacing="0" cellpadding="0">
			<tr>
				<td width="124" align="right" class="title_small">בונוס החודש</></td>
				<td width="8" align="left" bgcolor="#A3C312"><img src="images/info/arrow_green.gif" width="8" height="16"></td>
				<td width="38" align="center" bgcolor="#A3C312" class="value"><%=currentMonthRCount * OrderPrice%>₪</td>
			</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td width="100%" bgcolor="#F4F4F4" height="1"></td>
		</tr>
		<tr>
			<td width="100%" bgcolor="#F4F4F4" height="16" align="right"><table border="0" width="170" cellspacing="0" cellpadding="0">
			<tr>
				<td width="124" align="right" class="title_small">בונוס 10% הגברת מכירות</td>
				<td width="8" align="left" bgcolor="#BBBBBB"><img src="images/info/arrow_grey.gif" width="8" height="16"></td>
				<td width="38" align="center" bgcolor="#BBBBBB" class="value"><%=futureOrderMoney%>₪</td>
			</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td width="100%" bgcolor="#F4F4F4" height="2"></td>
		</tr>
		<tr>
			<td width="100%"></td>
		</tr>
		</table>
		</td>
		<td width="11"></td>
		<td valign="top" width="170"><table border="0" width="170" height="62" cellspacing="0" cellpadding="0">
		<tr>
			<td width="100%" background="images/info/title_bgr.gif" height="19" align="right" class="table_title">הזמנות של כל המחלקה&nbsp;&nbsp; </td>
		</tr>
		<tr>
			<td width="100%" bgcolor="#F4F4F4" height="2"></td>
		</tr>
		<tr>
			<td width="100%" bgcolor="#F4F4F4" height="16" align="right"><table border="0" width="170" cellspacing="0" cellpadding="0">
			<tr>
				<td width="124" align="right" class="title_small">הזמנות של כולם</td>
				<td width="8" align="left" bgcolor="#A3C312"><img src="images/info/arrow_green.gif" width="8" height="16"></td>
				<td width="38" align="center" bgcolor="#A3C312" class="value"><%=sumCurrentMonthRCount%></td>
			</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td width="100%" bgcolor="#F4F4F4" height="1"></td>
		</tr>
		<tr>
			<td width="100%" bgcolor="#F4F4F4" height="16" align="right"><table border="0" width="170" cellspacing="0" cellpadding="0">
			<tr>
				<td width="124" align="right" class="title_small">הזמנות לנציג המצטיין</td>
				<td width="8" align="left" bgcolor="#BBBBBB"><img src="images/info/arrow_grey.gif" width="8" height="16"></td>
				<td width="38" align="center" bgcolor="#BBBBBB" class="value"><%=maxCurrentMonthRCount%></td>
			</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td width="100%" bgcolor="#F4F4F4" height="2"></td>
		</tr>
		<tr>
			<td width="100%"></td>
		</tr>
		</table>
		</td>
		<td width="11"></td>
		<td valign="top" width="170"><table border="0" width="170" height="62" cellspacing="0" cellpadding="0">
		<tr>
			<td width="100%" background="images/info/title_bgr.gif" height="19" align="right" class="table_title">הזמנות חודש קודם&nbsp;&nbsp; </td>
		</tr>
		<tr>
			<td width="100%" bgcolor="#F4F4F4" height="2"></td>
		</tr>
		<tr>
			<td width="100%" bgcolor="#F4F4F4" height="16" align="right"><table border="0" width="170" cellspacing="0" cellpadding="0">
			<tr>
				<td width="124" align="right" class="title_small">יותר מחודש קודם</td>
				<td width="8" align="left" bgcolor="#A3C312"><img src="images/info/arrow_green.gif" width="8" height="16"></td>
				<td width="38" align="center" bgcolor="#A3C312" class="value"><%=IIF(pastMonthRCount > 0, Math.Round((currentMonthRCount/pastMonthRCount) * 100, 0), 0)%>%</td>
			</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td width="100%" bgcolor="#F4F4F4" height="1"></td>
		</tr>
		<tr>
			<td width="100%" bgcolor="#F4F4F4" height="16" align="right"><table border="0" width="170" cellspacing="0" cellpadding="0">
			<tr>
				<td width="124" align="right" class="title_small">הזמנות שלי חודש קודם</td>
				<td width="8" align="left" bgcolor="#BBBBBB"><img src="images/info/arrow_grey.gif" width="8" height="16"></td>
				<td width="38" align="center" bgcolor="#BBBBBB" class="value"><%=pastMonthRCount%></td>
			</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td width="100%" bgcolor="#F4F4F4" height="2"></td>
		</tr>
		<tr>
			<td width="100%"></td>
		</tr>
		</table>
		</td>
		<td width="11"></td>
		<td valign="top" width="170"><table border="0" width="170" height="62" cellspacing="0" cellpadding="0">
		<tr>
			<td width="100%" background="images/info/title_bgr.gif" height="19" align="right" class="table_title">הזמנות החודש&nbsp;&nbsp; </td>
		</tr>
		<tr>
			<td width="100%" bgcolor="#F4F4F4" height="2"></td>
		</tr>
		<tr>
			<td width="100%" bgcolor="#F4F4F4" height="16" align="right"><table border="0" width="170" cellspacing="0" cellpadding="0">
			<tr>
				<td width="124" align="right" class="title_small">הזמנות שלי החודש</td>
				<td width="8" align="left" bgcolor="#A3C312"><img src="images/info/arrow_green.gif" width="8" height="16"></td>
				<td width="38" align="center" bgcolor="#A3C312" class="value"><%=currentMonthRCount%></td>
			</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td width="100%" bgcolor="#F4F4F4" height="1"></td>
		</tr>
		<tr>
			<td width="100%" bgcolor="#F4F4F4" height="16" align="right"><table border="0" width="170" cellspacing="0" cellpadding="0">
			<tr>
				<td width="124" align="right" class="title_small">יעד חודשי</td>
				<td width="8" align="left" bgcolor="#BBBBBB"><img src="images/info/arrow_grey.gif" width="8" height="16"></td>
				<td width="38" align="center" bgcolor="#BBBBBB" class="value"><%=userTarget%></td>
			</tr>
			</table>
			</td>
		</tr>
		<tr>
			<td width="100%" bgcolor="#F4F4F4" height="2"></td>
		</tr>
		<tr>
			<td width="100%"><table border="0" width="100%" height="6" cellspacing="0" cellpadding="0">
			<tr>
				<td width="<%=100 - IIF(userTarget > 0, Math.Round((currentMonthRCount/userTarget) * 100, 0), 0)%>%" bgcolor="#D4D4D4"></td>
				<td width="<%=IIF(userTarget > 0, Math.Round((currentMonthRCount/userTarget) * 100, 0), 0)%>%" bgcolor="#A3C312"></td>
			</tr>
			</table>
			</td>
		</tr>
		</table>
		</td>
	</tr>
	</table>
		<!--table cellpadding="0" cellspacing="0" border="0" style="table-layout: fixed;" align="right">
		<tbody>
		<tr>
			<td rowspan="2" width="20"><a href="report_orders.aspx" target="_blank" title="דוח הזמנות לנציג"><img src="images/info/report_icon.gif" border="0" alt="דוח הזמנות לנציג"></a></td>
			<td width="45"><%=currentMonthRCount * OrderPrice%>₪</td>
			<th width="165">בונוס מתחילת החודש</th>			
			<td width="45" title="אחוז ביצוע לעומת החודש הקודם"><%=IIF(pastMonthRCount > 0, Math.Round((currentMonthRCount/pastMonthRCount) * 100, 0), 0)%>%</td>
			<th width="75">אחוז תפוקה</th>
			<td width="35"><%=pastMonthRCount%></td>
			<th width="125">ההזמנות בחודש הקודם</th>
			<td width="35"><%=currentMonthRCount%></td>
			<th width="125">הזמנות מתחילת החודש</th>
		</tr>
		<tr>
			<td><%=futureOrderMoney%>₪</td>
			<th>בונוס אם להגביר מכירות ב-10%</th>					
			<td title="סך כל ביצוע המחלקה מתחילת כל חודש עד לאותו יום נוכחי בחודש"><%=sumCurrentMonthRCount%></td>
			<th>סה"כ הזמנות</th>
			<td><%=maxCurrentMonthRCount%></td>
			<th>מספר ההזמנות max</th>
			<td><%=userTarget%></td>
			<th>יעד המינימאלי</th>
		</tr>		
		</tbody>
		</table-->
    </form>
  </body>
</html>