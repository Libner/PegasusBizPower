<%@ Page Language="vb" AutoEventWireup="false" Codebehind="gradesByDepId.aspx.vb" Inherits="bizpower_pegasus.gradesByDepId"%>
<!DOCTYPE HTML>
<html>
  <head>
    <title>gradesByDepId</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../IE4.css" rel="STYLESHEET" type="text/css">
	<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
	<script>
	function SendMailByCheck()
		{
	var fl = 0;
	var checkedValue = null; 
var cboxes = document.forms.Form1.chkTour;
  var len = cboxes.length;
    for (var i=0; i<len; i++) {
    if (cboxes[i].checked)
    {
    if (checkedValue == null)
    checkedValue=cboxes[i].value
    else
{
    checkedValue=checkedValue+','+ cboxes[i].value
 }   
      //  alert(i + (cboxes[i].checked?' checked ':' unchecked ') + cboxes[i].value);
      fl=1;
      }
    }
//alert(checkedValue)
	
	
	if (fl!=1)
 {
alert(" נא לסמן אנשי קשר על מנת לשלוח הודעה גורפת!");
	
		}
		else
		
{
  alert(" האם ברצונך לשלוח משובים עבור אנשי קשר המסומנים?");
	window.open ("SendMailFeedbackByContacts.aspx?sdep="+ escape(<%=DepartureId%>) +"&sCheck=" + escape(checkedValue),"SendMailByContacts","scrollbars=yes,menubar=yes,  width=710,height=600,left=10,top=10");
//window.open('SendMailFeedbackByContacts.aspx', 'SendMailByContacts', 'top=70, left=70, width=520, height=520');
	
	}
	
	}
	function SendSmsByCheck()
	{
	var fl = 0;
	var checkedValue = null; 
var cboxes = document.forms.Form1.chkTour;
  var len = cboxes.length;
    for (var i=0; i<len; i++) {
    if (cboxes[i].checked)
    {
    if (checkedValue == null)
    checkedValue=cboxes[i].value
    else
{
    checkedValue=checkedValue+','+ cboxes[i].value
 }   
      //  alert(i + (cboxes[i].checked?' checked ':' unchecked ') + cboxes[i].value);
      fl=1;
      }
    }
//alert(checkedValue)
	
	
	if (fl!=1)
 {
alert(" נא לסמן אנשי קשר על מנת לשלוח הודעה גורפת!");
	
		}
		else
		
{
  alert(" האם ברצונך לשלוח משובים עבור אנשי קשר המסומנים?");
	window.open ("SendSMSFeedbackByContacts.aspx?"+"sdep="+ escape(<%=DepartureId%>) +"&sCheck=" + escape(checkedValue),"SendSmsByContacts","scrollbars=yes,menubar=yes, width=710,height=600,left=10,top=10");

	
	}
	
	}
	function cball_onclick(source) 
	{

  checkboxes = document.getElementsByName('chkTour');
  for(var i=0, n=checkboxes.length;i<n;i++) {
    checkboxes[i].checked = source.checked;
  }
}
 
function FuncSort(par,srt)
{



var r  ;
r="&"+par +"=ASC"
//query = query.replace(r, "");
r="&"+par +"=DESC"
//alert (srt)

window.location ="gradesByDepId.aspx?depId="+<%=DepartureId%>+"&"+par +"="+ srt


}
function openSendReport()
		{
		var query 
		query="<%=qrystring%>"
if (query!="")
{
			window.open("https://www.pegasusisrael.co.il/biz_form/SendreportFeedbackScreen.aspx?"+query,"ReportExcel","scrollbars=yes,menubar=yes, toolbar=yes, width=710,height=600,left=10,top=10");

		}
if (query=="")
{

window.open ("https://www.pegasusisrael.co.il/biz_form/SendreportFeedbackScreen.aspx?depId="+<%=DepartureId%>,"ReportExcel","scrollbars=yes,menubar=yes, toolbar=yes, width=710,height=600,left=10,top=10");
//http://www.pegasusisrael.co.il/biz_form
}
		

		}


		function openPdfReport()
		{
			var query 
		query="<%=qrystring%>"
		//alert(query)
if (query!="")
{
			window.open("https://www.pegasusisrael.co.il/biz_form/reportFeedbackScreenPdf.aspx?"+query,"ReportPdf","scrollbars=yes,menubar=yes, toolbar=yes, width=710,height=600,left=10,top=10");

		}
if (query=="")
{

window.open ("https://www.pegasusisrael.co.il/biz_form/reportFeedbackScreenPdf.aspx?depId="+<%=DepartureId%>,"ReportPdf","scrollbars=yes,menubar=yes, toolbar=yes, width=710,height=600,left=10,top=10");

}
		}
</script>
  </head>
  <body style="margin:0px">

    <form id="Form1" method="post" runat="server">
    <table border=0 cellpadding=0 cellspacing=0 width=100% align=left>
  
   <tr>
<td width=100% style="color:#000000;background-color:#FFD011;"> 
<table border=0 cellpadding=0 cellpadding=0 width=100%>
<tr><td align=left width =100><a class="button_edit_1" style="width:94;" href='javascript:void(0)' onclick="return openPdfReport();">הצג דוח PDF</a></td>
	<td align=left width =100><a class="button_edit_1" style="width:94;" href='javascript:void(0)' onclick="return openSendReport();">שליחת מייל</a></td>
	<td align=center style="color:#000000;background-color:#FFD011;font-size:14px;font-weight:bold">מסך ציוני  <%=TourName%></td>
   </tr>

</table></td>
</tr>
<tr><td>
<table border=0 cellpadding=0 cellspacing=0 width=100% align=left>
<tr bgcolor=#d8d8d8 style="height:25px" >
	<td class="title_sort"  align="center">כמות האנשים הרשומים לטיול</td>	
	<td class="title_sort"  align="center">כמות האנשים שמילאו משוב</td>	
	<td class="title_sort"  align="center">ציון הטיול הכללי</td>	
	<td class="title_sort"  align="center">שם המדריך</td>	
	<td class="title_sort"  align="center">קוד הטיול</td>	
	</tr>
	<tr  style="height:25px" >
	<td align="center"><%=CountMembers%></td>	
	<td   align="center"><%=CountFeedBack%></td>	
	<th   align="center" <%if CountFeedBack>0 then%><% if Math.Round(allTourGrade/CountFeedBack,2)<=70 then%>style="background-color:#ff0000" <%else if Math.Round(allTourGrade/CountFeedBack,2)>70 and Math.Round(allTourGrade/CountFeedBack,2)<=80 then%>style="background-color:#FFFF00" <%else if Math.Round(allTourGrade/CountFeedBack,2)>80 and Math.Round(allTourGrade/CountFeedBack,2)<=90 then%>style="background-color:#00FF00" <%else if Math.Round(allTourGrade/CountFeedBack,2)>90 then%>style="background-color:#1F7246"<%end if%>>
	<%if allTourGrade>0 then%>		<%=Math.Round(allTourGrade/CountFeedBack,2)%>%<%end if%><%end if%></th>	
	<th   align="center" class="title_show_form"><%=GuideName%></th>	
	<th  align="center" class="title_show_form"><%=DepartureCode%></th>	
	</tr>
	</table>
<br><br clear=all><br><br clear=all>
   <table border=0 cellpadding=0 cellspacing=0 width=100% align=left>
   <tr><td align=center style="color:#000000;background-color:#FFD011;font-size:14px;font-weight:bold">אנשי קשר אשר מילאו חוות דעת אלקטרוניות </td></tr>
   </table><br clear=all>
<table border=0 cellpadding=1 cellspacing=1 width=100% align=left>
<tr bgcolor=#d8d8d8 style="height:25px" >
	<th align="right" class="title_show_form"  width =10%>&nbsp;<a href="javascript:FuncSort('sort_6','ASC')"><img src="../../../images/arrow_top<%If Request.Querystring("sort_6")="ASC" then %>_act<%end if%>.gif"  title="למיין לפי טול" border=0></a><a href="javascript:FuncSort('sort_6','DESC')"><img src="../../../images/arrow_bot<%If Request.Querystring("sort_6")="DESC" then %>_act<%end if%>.gif"  title="למיין לפי טיול" border=0></a>&nbsp;&nbsp;ציון כללי של המשוב&nbsp;</th>	
		<th  align="right" class="title_show_form"  width =10%>&nbsp;<a href="javascript:FuncSort('sort_5','ASC')"><img src="../../../images/arrow_top<%If Request.Querystring("sort_5")="ASC" then %>_act<%end if%>.gif"  title="למיין לפי טול" border=0></a><a href="javascript:FuncSort('sort_5','DESC')"><img src="../../../images/arrow_bot<%If Request.Querystring("sort_5")="DESC" then %>_act<%end if%>.gif"  title="למיין לפי טיול" border=0></a>&nbsp;&nbsp;קטגוריות הציון השני <br>הכי גרוע שהלקוח נתן&nbsp;</th>	
	<th  align="right" class="title_show_form"  width =19%><a href="javascript:FuncSort('sort_4','ASC')"><img src="../../../images/arrow_top<%If Request.Querystring("sort_4")="ASC" then %>_act<%end if%>.gif"  title="למיין לפי טול" border=0></a><a href="javascript:FuncSort('sort_4','DESC')"><img src="../../../images/arrow_bot<%If Request.Querystring("sort_4")="DESC" then %>_act<%end if%>.gif"  title="למיין לפי טיול" border=0></a>&nbsp;&nbsp; קטגוריות הציון הכי גרוע שהלקוח נתן&nbsp;</th>	
	<th  align="right" class="title_show_form"  width =13%><a href="javascript:FuncSort('sort_3','ASC')"><img src="../../../images/arrow_top<%If Request.Querystring("sort_3")="ASC" then %>_act<%end if%>.gif"  title="למיין לפי טול" border=0></a><a href="javascript:FuncSort('sort_3','DESC')"><img src="../../../images/arrow_bot<%If Request.Querystring("sort_3")="DESC" then %>_act<%end if%>.gif"  title="למיין לפי טיול" border=0></a>&nbsp;&nbsp;שביעות רצונך הכללית מהטיול&nbsp;</th>	
	<th  align="right" class="title_show_form"  width =14%><a href="javascript:FuncSort('sort_2','ASC')"><img src="../../../images/arrow_top<%If Request.Querystring("sort_2")="ASC" then %>_act<%end if%>.gif"  title="למיין לפי טול" border=0></a><a href="javascript:FuncSort('sort_2','DESC')"><img src="../../../images/arrow_bot<%If Request.Querystring("sort_2")="DESC" then %>_act<%end if%>.gif"  title="למיין לפי טיול" border=0></a>&nbsp;&nbsp;בהנחה ויתאפשר, באיזה מידה סביר כי תשוב לטייל עם פגסוס&nbsp;</th>	
	<th align="right" class="title_show_form" width=6% dir=rtl>מאיזה הודעה <BR>הלקוח פתח את המשוב</th>

	<th  align="right" class="title_show_form" width =26%><a href="javascript:FuncSort('sort_1','ASC')"><img src="../../../images/arrow_top<%If Request.Querystring("sort_1")="ASC" then %>_act<%end if%>.gif"  title="למיין לפי טול" border=0></a><a href="javascript:FuncSort('sort_1','DESC')"><img src="../../../images/arrow_bot<%If Request.Querystring("sort_1")="DESC" then %>_act<%end if%>.gif"  title="למיין לפי טיול" border=0></a>&nbsp;איש קשר&nbsp;&nbsp;</th>	
		<th  align="right" class="title_show_form" width =2%>&nbsp;</th>	
	</tr>
	<asp:Repeater ID=rptCustomers Runat=server>
	<ItemTemplate>
<tr style="background-color: rgb(201, 201, 201);height:30px;"  onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#C9C9C9';" >
	<td align=center><a style="text-decoration:none;color:#000000;" href="javascript:void(0)" onclick="window.open('Feedback.aspx?status=2&depId=<%=DepartureId%>&companyId=<%#Container.DataItem("Company_Id")%>&contactId=<%#Container.DataItem("Contact_Id")%>','AddCat','top=10,left=10,width=900,height=450,scrollbars=1')" class="button_admin_1"><%#Container.DataItem("TourGrade")%>%
	<asp:Label ID="lblTourGrade" name="lblTourGrade" Runat=server></asp:Label></td>
	<td align=center><%#Container.Dataitem("NextMinGrade")%><%'#func.Feedback_GetNextMinGrade(DepartureId,Container.DataItem("Contact_Id"))%></td>
	<td align=center><%#Container.DataItem("FaqGradeMinValue")%></td>
	<td align=center><%#Container.DataItem("FAQ14")%></td>
	<td align=center><%#Container.DataItem("FAQ13")%></td>
	<td align=center><%#Container.DataItem("SourceDetection")%></td>
<td align="right"><a  target="_parent"  style="text-decoration:none;color:#000000;" href="../companies/contact.asp?companyId=<%#Container.DataItem("Company_Id")%>&contactId=<%#Container.DataItem("Contact_Id")%>"><%#Container.DataItem("CONTACT_NAME")%></a>&nbsp;</td>
<td align=center>&nbsp;<%#Container.ItemIndex+1%></td>
</tr>
	</ItemTemplate>
	<AlternatingItemTemplate>
<tr bgcolor="#E6E6E6" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#E6E6E6';" style="background-color: rgb(230, 230, 230);height:25px;" >
	<td align=center><a style="text-decoration:none;color:#000000;" href="javascript:void(0)" onclick="window.open('Feedback.aspx?status=2&depId=<%=DepartureId%>&companyId=<%#Container.DataItem("Company_Id")%>&contactId=<%#Container.DataItem("Contact_Id")%>','AddCat','top=10,left=10,width=900,height=450,scrollbars=1')" class="button_admin_1"><%#Container.DataItem("TourGrade")%>%
	<asp:Label ID="lblTourGrade" name="lblTourGrade" Runat=server></asp:Label></td>
	<td align=center><%#Container.Dataitem("NextMinGrade")%><%'#func.Feedback_GetNextMinGrade(DepartureId,Container.DataItem("Contact_Id"))%></td>
	<td align=center><%#Container.DataItem("FaqGradeMinValue")%></td>
	<td align=center><%#Container.DataItem("FAQ14")%></td>
	<td align=center><%#Container.DataItem("FAQ13")%></td>
	<td align=center><%#Container.DataItem("SourceDetection")%></td>

<td align="right"><a  target="_parent"  style="text-decoration:none;color:#000000;" href="../companies/contact.asp?companyId=<%#Container.DataItem("Company_Id")%>&contactId=<%#Container.DataItem("Contact_Id")%>"><%#Container.DataItem("CONTACT_NAME")%></a>&nbsp;<%#IIF(func.GetContactByDepBitulim(DepartureId,Container.DataItem("Contact_Id"))="1","<span style=background-color:#ff0000;color:#ffffff>&nbsp;בוטל</span>","")%></td>
<td align=center>&nbsp;<%#Container.ItemIndex+1%></td>
</tr>

	</AlternatingItemTemplate>
	
	</asp:Repeater>
</table>
</td>
</tr></table>
<br clear=all>
<br clear=all>
<br clear=all>
	<asp:Repeater ID="rptCustomers2" Runat=server>
<HeaderTemplate>
   <table border=0 cellpadding=0 cellspacing=0 width=100% align=left>
   <tr>
   <td align=left style="background-color:#FFD011;padding-left:2px"><%if request.Cookies("bizpegasus")("UserSendSmsGeneralScreen")="1" then%>
   <a class="button_edit_1" style="font-size:16px color:#ff0000;width:100px"  onclick="javascript:SendMailByCheck()" style="width: 68px;background-color:#F7CD65;color:#000000">שלח משוב במייל</a>
   <a class="button_edit_1" style="font-size:16px color:#ff0000;width:100px"  onclick="javascript:SendSmsByCheck()" style="width: 68px;background-color:#F7CD65;color:#000000">שלח משוב בסמס</a>
   <%end if%>
   </td>
 
   <td align=center style="color:#000000;background-color:#FFD011;font-size:14px;font-weight:bold">אנשי קשר אשר לא מילאו חוות דעת אלקטרוניות</td>
    </tr>
   </table><br clear=all>
   <table border=0 cellpadding=1 cellspacing=1 width=100% align=left>
<tr bgcolor=#d8d8d8 style="height:25px" >
	<th align="center" class="title_show_form"  width =10%>
	בחר הכל<br>
	<input type="checkbox" language="javascript" onclick="return cball_onclick(this)" title="הכל" id="cb_all" name="cb_all"></th>	
		<th  align="right" class="title_show_form"  width =10%>מספר הפעמים שנשלחה ההודעה ללקוח באמצעי כלשהו</th>	
	<th  align="right" class="title_show_form"  width =10%>הלקוח לחץ על הקישור להתחלת המשוב אך יצא ממנו ללא התחלת מילוי</th>	
	<th  align="right" class="title_show_form"  width =10% dir=rtl>האם נשלחה הודעת מייל?</th>	
	<th  align="right" class="title_show_form"  width =10% dir=rtl>האם נשלחה הודעת סמס?</th>	
	<th  align="right" class="title_show_form" width =48%>איש קשר&nbsp;</th>	
	<th  align="right" class="title_show_form" width =2%>&nbsp;</th>	
	</tr>
   </HeaderTemplate>
   
<ItemTemplate>
<tr style="height:30px;background-color: rgb(201, 201, 201);")" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#C9C9C9';" >
	<td align=center><%if request.Cookies("bizpegasus")("UserSendSmsGeneralScreen")="1" then%>
	<input type="checkbox" ID="chkTour"  NAME="chkTour" value="<%#Container.DataItem("Contact_Id")%>">&nbsp;<%end if%>
</td>
	<td align=center><asp:label ID="lblCountSendAll" Runat=server></asp:label><%'#func.dbNullFix(Container.DataItem("SendCount"))%></td>
	<td align=center><%#IIF (func.dbNullFix(Container.DataItem("OpenDevice"))<>"","כן" ,"לא")%></td>
<td align=center><asp:Label ID=lblSendmail Runat=server></asp:Label></td>
	<td align=center><%#IIF (func.dbNullFix(Container.DataItem("date_send"))<>"","כן" ,"לא" )%></td>
<td align="right"><a  target="_parent"  style="text-decoration:none;color:#000000;" href="../companies/contact.asp?companyId=<%#Container.DataItem("Company_Id")%>&contactId=<%#Container.DataItem("Contact_Id")%>"><%#Container.DataItem("CONTACT_NAME")%></a>&nbsp;<%#IIF(func.GetContactByDepBitulim(DepartureId,Container.DataItem("Contact_Id"))="1","<span style=background-color:#ff0000;color:#ffffff>&nbsp;בוטל</span>","")%></td>
<td align=center nowrap>&nbsp;<%#Container.ItemIndex+1%></td>
</tr>

</ItemTemplate>
<FooterTemplate></table></FooterTemplate>
</asp:Repeater>
<br clear=all>
<br clear=all>
<br clear=all>
	<asp:Repeater ID="rptCustomers3" Runat=server>
<HeaderTemplate>
   <table border=0 cellpadding=0 cellspacing=0 width=100% align=left>
   <tr><td align=center style="color:#000000;background-color:#FFD011;font-size:14px;font-weight:bold">אנשי קשר שמילאו משוב באופן חלקי</td></tr>
   </table><br clear=all>
   <table border=0 cellpadding=1 cellspacing=1 width=100% align=left>
<tr bgcolor=#d8d8d8 style="height:25px" >
	<th align="center" class="title_show_form"  width =20% colspan=2>שלח שוב בקשה ללקוח למילוי המשוב&nbsp;</th>	

	<th>&nbsp;</th>
		<th  align="right" class="title_show_form"  width =10%>מספר הפעמים שנשלחה ההודעה ללקוח באמצעי כלשהו</th>	
	<th  align="right" class="title_show_form"  width =10% dir=rtl>האם פתח הלינק דרך טאבלט או דרך מחשב ביתי?</th>	
	<th  align="right" class="title_show_form"  width =10% dir=rtl>באיזה חלק של המשוב הפסיק הלקוח?</th>	
		<th align="right" class="title_show_form" width=6% dir=rtl>מאיזה הודעה <BR>הלקוח פתח את המשוב</th>
	<th  align="right" class="title_show_form" width =38%>איש קשר&nbsp;</th>	
	<th  align="right" class="title_show_form" width =2%>&nbsp;</th>	
	</tr>
   </HeaderTemplate>
   
<ItemTemplate>
<tr style="background-color: rgb(201, 201, 201);height:30px;"  onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#C9C9C9';" >
	<td align=center  colspan=2><%if Request.Cookies("bizpegasus")("UserSendSmsGeneralScreen")="1" then%>
	<a class="button_edit_1" href="javascript:void window.open('SendMailFeedBack.aspx?DepartureId=<%=DepartureId%>&companyId=<%#Container.DataItem("Company_Id")%>&contactId=<%#Container.DataItem("Contact_Id")%>','SendMailFeedback','top=70, left=70, width=520, height=520, scrollbars=1');" onclick="return window.confirm('?האם ברצונך לשלוח מייל לאיש קשר');" style="width: 120px;background-color:#F7CD65;color:#000000">שלח משוב במייל</a>
	<a class="button_edit_1" href="javascript:void window.open('SendSMSFeedBack.aspx?DepartureId=<%=DepartureId%>&companyId=<%#Container.DataItem("Company_Id")%>&contactId=<%#Container.DataItem("Contact_Id")%>','SendSMSFeedback','top=70, left=70, width=520, height=520, scrollbars=1');" onclick="return window.confirm('?האם ברצונך לשלוח סמס לאיש קשר');" style="width: 120px;background-color:#F7CD65;color:#000000">שלח משוב בסמס</a>
	<%end if%>
</td>
<td align=center><a class="button_edit_1" style="width: 120px;background-color:#736BA6;color:#ffffff" href="javascript:void(0)" onclick="window.open('Feedback.aspx?status=1&depId=<%=DepartureId%>&companyId=<%#Container.DataItem("Company_Id")%>&contactId=<%#Container.DataItem("Contact_Id")%>','AddCat','top=10,left=10,width=900,height=450,scrollbars=1')" class="button_admin_1">הצג משוב חלקי</a></td>
	<td align=center><asp:label ID="lblCountSendAll" Runat=server></asp:label><%'#func.dbNullFix(Container.DataItem("SendCount"))%></td>

<td align=center><%#Container.DataItem("OpenDevice")%></td>
<td align=center><%#Container.DataItem("LastFAQ")%> שאלה<%'#func.GetFeedBackCategoryName(Container.DataItem("LastFAQ"))%></td>
	<td align=center><%#Container.DataItem("SourceDetection")%></td>

<td align="right"><a  target="_parent"  style="text-decoration:none;color:#000000;" href="../companies/contact.asp?companyId=<%#Container.DataItem("Company_Id")%>&contactId=<%#Container.DataItem("Contact_Id")%>">
<%#Container.DataItem("CONTACT_NAME")%></a>&nbsp;<%#IIF(func.GetContactByDepBitulim(DepartureId,Container.DataItem("Contact_Id"))="1","<span style=background-color:#ff0000;color:#ffffff>&nbsp;בוטל</span>","")%></td>
<td align=center><%#Container.ItemIndex+1%></td>
</tr>

</ItemTemplate>
<FooterTemplate></table></FooterTemplate>
</asp:Repeater>
    </form>

  </body>
</html>
