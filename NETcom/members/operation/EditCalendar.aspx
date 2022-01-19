<%@ Page Language="vb" AutoEventWireup="false" Codebehind="EditCalendar.aspx.vb" Inherits="bizpower_pegasus.EditCalendar" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML lang="en">
	<HEAD>
		<title>
			<%=fTitle%>
		</title>
		<meta charset="utf-8">
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
			<link rel="stylesheet" href="/resources/demos/style.css">
				<script src="https://code.jquery.com/jquery-1.12.4.js"></script>
				<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
				<script>
  $( function() {
    $("#datepicker").datepicker({
                 onSelect:function(selectedDate){
                 document.getElementById("datepickerN").value=selectedDate
           document.forms[0].submit()
                       }
                       }

    );
    $.datepicker.setDefaults($.datepicker.regional["he"]);
      $.datepicker.setDefaults({ dateFormat: 'dd/mm/yy' });
    $("#datepicker").datepicker('setDate', '<%=fValue%>');


  } );
 
function DeleteCalendar()
{
document.getElementById("datepickerN").value=""
document.forms[0].submit()
}
				</script>
				<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
	</HEAD>
	<body style="MARGIN:0px" onload="self.focus();">
		<center>
			<form id="Form1" method="post" runat="server">
				<table width="100%" align="center" cellspacing="0" cellpadding="2" border="0">
					<tr>
						<td bgcolor="#e6e6e6" height="5" align="center" width="100%"><b><%=fTitle%><br>
								<%=DepartureCode%>
							</b>
						</td>
					</tr>
					<tr>
						<td align="center" class="td_admin_4" height="10"></td>
					</tr>
					<tr>
						<td align="center" class="td_admin_4">
							<div><input id="datepickerN" name="datepickerN" type="hidden"></div>
							<br>
							<div id="datepicker"></div>
						</td>
					</tr>
					<tr><td align=center><a href="#" onClick=DeleteCalendar();><img src="../../images/delete_icon.gif" border="0" >מחיקה תאריך</a></td></tr>
					<tr>
						<td align="center" class="td_admin_4" height="10"></td>
					</tr>
				
				</table>
			</form>
		</center>
	</body>
</HTML>
