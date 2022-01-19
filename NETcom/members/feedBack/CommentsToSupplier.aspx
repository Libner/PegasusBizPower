<%@ Page Language="vb" AutoEventWireup="false" EnableViewState="true" Codebehind="CommentsToSupplier.aspx.vb" Inherits="bizpower_pegasus2018.CommentsToSupplier" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>CommentsToSupplier</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<script>
    function OpenGoogleTranslate()
    {
    //https://translate.google.co.il/?hl=iw#iw/en/
    }
		</script>
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
		<script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.1/jquery.min.js"></script>
	<script>
	function openUpdateFeeadback19(id)
	{
	//alert(id)
	h = 500;
		w = 800;
		S_Wind = window.open("UpdateFeeadback19.aspx?id=" + id+"&sdep=<%=DepartureId%>", "S_Wind", "scrollbars=1,toolbar=0,top=50,left=50,width=" + w + ",height=" + h + ",align=center,resizable=1");
		S_Wind.focus();	
		return false;
	}
	function openAddComments()
	{
	h = 500;
		w = 800;
		S_Wind = window.open("AddComments.aspx?sdep=<%=DepartureId%>", "S_Wind", "scrollbars=1,toolbar=0,top=50,left=50,width=" + w + ",height=" + h + ",align=center,resizable=1");
		S_Wind.focus();	
		return false;

	}
	</script>
		<script>
	function checkDelete()
	{

		
		
		return window.confirm("?האם ברצונך למחוק נתונים");
		
	}
	    function CheckForm()
        {
        
              if (document.getElementById("ContentText").value == '') 
            {
                  alert("אנא הכנס  הערות");
                  document.getElementById("ContentText").focus()
                return false;
            }
       //    document.getElementById("Form1").submit();
            return true;

        }
		</script>
		<script>
$(document).ready(function () {

  $('.clsAlphaNoOnly').keypress(function (e) {  // Accept only alpha numerics, no special characters 
 // alert(e.which)
  if (e.which>=1488 &&e.which<=1514)
  {
    alert("אנא הכנס טקסט באנגלית בלבד ")
  return false;
  }
      /*  var regex = new RegExp("^[a-zA-Z0-9.,@%*()=-]+$");
        var str = String.fromCharCode(!e.charCode ? e.which : e.charCode);
        if (regex.test(str)) {
            return true;
        }*/

     /*   e.preventDefault();
        alert("אנא הכנס טקסט באנגלית בלבד ")
        return false;*/
    }); 
})
		</script>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<center>
			<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
				<tr>
					<td height="20" colspan="3">&nbsp;</td>
				</tr>
				<tr>
					<td><%if false then%><input type="button" name="GButton" id="GButton" style="WIDTH:120px" value="Google translate"
							class="but_menu" onclick="window.open('https://translate.google.co.il/?hl=iw#iw/en/','winUPD','top=20, left=10, width=1000, height=350, scrollbars=1');"><%end if%></td>
				<td width=120>&nbsp;</td>
				<td align=left class="title_show_form" nowrap><a class="button_edit_1" style="width:120;" href="javascript:void(0)" onclick="return openAddComments();">הוספת הערות ידניות</a>&nbsp;
				</td>
					<td height="25" align="center" width="100%" dir="rtl" style="FONT-SIZE:14pt" nowrap><b>הוספת 
							הערות חשובות לטבלת המשובים באזור האישי של הספק בגין טיול
							<%=func.GetDepatureCode(DepartureId)%>
							<br>
						</b>
					</td>
				</tr>
				<tr>
					<td align="center" colspan="4">
						<table cellpadding="1" cellspacing="1" width="100%" align="center">
							<tr>
								<td height="20" colspan="4">&nbsp;</td>
							</tr>
							<asp:Repeater ID="rptCustomers" Runat="server">
								<HeaderTemplate>
									<tr bgcolor="#d8d8d8" style="height:25px">
									<td class="title_show_form" >&nbsp;</td>
										<td class="title_show_form" >&nbsp;</td>
										<td align="center" width="75%">
										<table border=0 cellpadding=0 cellspacing=0 width=100%>
										<tr>
										<th align="center" class="title_show_form" width=100% >&nbsp;&nbsp;הערות&nbsp;&nbsp;</th></tR>
										</table></td>
										
										
										<th align="center" class="title_show_form" width="10%" nowrap>
											&nbsp;&nbsp;שם משתמש &nbsp;&nbsp;</th>
										<th align="center" class="title_show_form" width="10%" nowrap>
											&nbsp;&nbsp; תאריך &nbsp;&nbsp;</th>
										<th align="center" class="title_show_form" width="5%" nowrap>
											&nbsp;&nbsp; ספק&nbsp;&nbsp;</th>
										<th  class="title_show_form" >&nbsp;</th>	
										
									</tr>
								</HeaderTemplate>
								<ItemTemplate>
									<tr style="background-color: rgb(201, 201, 201);height:30px;" onmouseover="this.style.backgroundColor='#B1B0CF';"
										onmouseout="this.style.backgroundColor='#C9C9C9';">
										<td><a href="CommentsToSupplier.aspx?sdep=<%=DepartureId%>&deleteId=<%#Container.DataItem("Id")%>" onclick="return checkDelete()"><img src="../../images/delete_icon.gif" border="0"></a></td>
										<td><a href="" onclick="return openUpdateFeeadback19('<%#Container.DataItem("Id")%>')" target="_blank"><img src="../../images/edit_icon.gif" border="0" alt="עדכון"></a></td>
										<td align="left" style=padding-left:3px><%#Container.DataItem("Content_Text")%></td>
										<td nowrap align="center"><%#Container.DataItem("UserName")%></td>
										<td nowrap align="center"><%#Container.DataItem("Date")%></td>
										<td  align="center"><%#Container.DataItem("supplier_Name")%></td>
										<td  align="center"><%#Container.ItemIndex+1%></td>
									</tr>
								</ItemTemplate>
								<AlternatingItemTemplate>
									
									<tr bgcolor="#E6E6E6" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#E6E6E6';"
										style="background-color: rgb(230, 230, 230);height:25px;">
										<td><a href="CommentsToSupplier.aspx?sdep=<%=DepartureId%>&deleteId=<%#Container.DataItem("Id")%>" onclick="return checkDelete()"><img src="../../images/delete_icon.gif" border="0"></a></td>
									
									<td><a href="" onclick="return openUpdateFeeadback19('<%#Container.DataItem("Id")%>')"><img src="../../images/edit_icon.gif" border="0" alt="עדכון"></a></td>
										<td  align="left" style=padding-left:3px ><%#Container.DataItem("Content_Text")%></td>
										<td nowrap align="center"><%#Container.DataItem("UserName")%></td>
										<td nowrap align="center"><%#Container.DataItem("Date")%></td>
										<td  align="center"><%#Container.DataItem("supplier_Name")%></td>
										<td  align="center"><%#Container.ItemIndex+1%></td>
										</tr>
								</AlternatingItemTemplate>
							</asp:Repeater>
						</table>
						<%if false then 'add new comments---%>
						<CENTER></CENTER>
						<br>
						<form id="Form1" method="post" runat="server">
							<table width="100%" align="center" cellspacing="0" cellpadding="2" border="0">
								<tr>
									<td height="20">&nbsp;</td>
								</tr>
								<tr>
									<td width="20%" align="center" valign="top">
										<table border="0" cellpadding="5" cellspacing="5">
											<tr>
												<td align="center" class="td_admin_4" height="10"></td>
											</tr>
											<tr>
												<td align="right"><table border="0">
														<tr>
															<td align="right">
																<select id="fselect" name="fselect" dir="rtl" runat="server" style="FONT-SIZE:9pt;FONT-FAMILY:arial;WIDTH:300px">
																</select>
															</td>
														</tr>
													</table>
												</td>
											</tr>
											<tr>
												<td align="right" class="td_admin_4"><input name="dateP" id="dateP" dir="ltr" runat="server" disabled></td>
												<td align="right">&nbsp;תאריך&nbsp;</td>
											</tr>
											<tr>
												<td align="right" class="td_admin_4"><input name="UserName" id="UserName" dir="ltr" runat="server" disabled></td>
												<td align="right">&nbsp;שם משתמש&nbsp;</td>
											</tr>
											<tr>
												<td align="right" class="td_admin_4"><textarea id="ContentText" name="ContentText" dir="ltr" style="HEIGHT:40px;WIDTH:450px;LINE-HEIGHT:120%"
														class="clsAlphaNoOnly"></textarea></td>
												<td align="right">&nbsp;הערות מהמשוב</td>
											</tr>
											<tr>
												<td height="20">&nbsp;</td>
											</tr>
											<tr>
												<td align="center"><table border="0" cellpadding="0" cellspacing="0" align="center">
														<tr>
															<td align="center" class="td_admin_4">
																<input type="submit" name="Button1" style="WIDTH:300px" value="טען הערות משובים לאזור האישי של הספק/ספקים"
																	id="Button1" class="but_menu" onClick="return CheckForm()">
															</td>
														</tr>
													</table>
												</td>
											</tr>
										</table>
									</td>
								</tr>
							</table>
					</td>
				</tr>
			</table>
			</FORM>
			<%end if%>
		</center>
	</body>
</HTML>
