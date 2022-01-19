<%@ Page Language="vb" AutoEventWireup="false" Codebehind="AddComments.aspx.vb" Inherits="bizpower_pegasus2018.AddComments"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
    <title>AddComments</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
 	<script>
    function OpenGoogleTranslate()
    {
    //https://translate.google.co.il/?hl=iw#iw/en/
    }
		</script>
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
		<script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.1/jquery.min.js"></script>
	<script>
  
	
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
  </head>
 <body MS_POSITIONING="GridLayout">
		<center>
			<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
				<tr>
					<td height="20" colspan="3">&nbsp;</td>
				</tr>
				<tr>
					<td><input type="button" name="GButton" id="GButton" style="WIDTH:120px" value="Google translate"
							class="but_menu" onclick="window.open('https://translate.google.co.il/?hl=iw#iw/en/','winUPD','top=20, left=10, width=1000, height=350, scrollbars=1');"></td>
					<td width="120">&nbsp;</td>
					<td height="25" align="center" width="100%" dir="rtl" style="FONT-SIZE:14pt" nowrap><b>הוספת 
							הערות חשובות לטבלת המשובים באזור האישי של הספק בגין טיול
							<%=func.GetDepatureCode(DepartureId)%>
							<br>
						</b>
					</td>
				</tr>
				<tr>
					<td align="center" colspan="3">
						<table cellpadding="1" cellspacing="1" width="100%" align="center">
							<tr>
								<td height="20" colspan="4">&nbsp;</td>
							</tr>
						
						</table>
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
												<td align="right" class="td_admin_4"><textarea id="ContentText" name="ContentText" dir="ltr" style="HEIGHT:80px;WIDTH:450px;LINE-HEIGHT:120%"
														class="clsAlphaNoOnly"></textarea></td>
												<td align="right">&nbsp;הערות ידניות</td>
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
		</center>
	</body>
</HTML>
