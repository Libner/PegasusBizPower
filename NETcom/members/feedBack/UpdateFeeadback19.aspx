<%@ Page Language="vb" AutoEventWireup="false" Codebehind="UpdateFeeadback19.aspx.vb" Inherits="bizpower_pegasus2018.UpdateFeeadback19"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
    <title>UpdateFeeadback19</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    	<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
		<script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.1/jquery.min.js"></script>
		<script>
  
	
	    function CheckForm()
        {
        
              if (document.getElementById("ContentText").value == '') 
            {
                  alert("��� ����  �����");
                  document.getElementById("ContentText").focus()
                return false;
            }
       document.getElementById("Form1").submit();
            return true;

        }
		</script>
		<script>
$(document).ready(function () {

  $('.clsAlphaNoOnly').keypress(function (e) {  // Accept only alpha numerics, no special characters 
 // alert(e.which)
  if (e.which>=1488 &&e.which<=1514)
  {
    alert("��� ���� ���� ������� ���� ")
  return false;
  }
      /*  var regex = new RegExp("^[a-zA-Z0-9.,@%*()=-]+$");
        var str = String.fromCharCode(!e.charCode ? e.which : e.charCode);
        if (regex.test(str)) {
            return true;
        }*/

     /*   e.preventDefault();
        alert("��� ���� ���� ������� ���� ")
        return false;*/
    }); 
})
		</script>
  </head>
  <body MS_POSITIONING="GridLayout">

    <form id="Form1" method="post" runat="server">
		<table border="0" cellpadding="0" cellspacing="0" width="100%" align="center">
				<tr>
					<td height="20" colspan="3">&nbsp;</td>
				</tr>
				<tr>
					<td><input type="button" name="GButton" id="GButton" style="WIDTH:120px" value="Google translate"
							class="but_menu" onclick="window.open('https://translate.google.co.il/?hl=iw#iw/en/','winUPD','top=20, left=10, width=1000, height=350, scrollbars=1');"></td>
					<td width="120">&nbsp;</td>
					<td height="25" align="center" width="100%" dir="rtl" style="FONT-SIZE:14pt" nowrap><b>����� 
							����� ������ ����� ������� ����� ����� �� ���� ���� ����
							<%=func.GetDepatureCode(DepartureId)%>
							<br>
						</b>
					</td>
				</tr></table><br>
	<table border="0" cellpadding="5" cellspacing="5">
											<tr>
												<td align="center" class="td_admin_4" height="10"></td>
											</tr>
											<tr>
												<td align="right"><table border="0">
														<tr>
															<td align="right">�� ������</td>
														</tr>
													</table>
												</td>
											</tr>
											<tr>
												<td align="right" class="td_admin_4"><input name="dateP" id="dateP" dir="ltr" runat="server" disabled></td>
												<td align="right">&nbsp;�����&nbsp;</td>
											</tr>
											<tr>
												<td align="right" class="td_admin_4"><input name="UserName" id="UserName" dir="ltr" runat="server" disabled></td>
												<td align="right">&nbsp;�� �����&nbsp;</td>
											</tr>
												<tr>
												<td align="right" class="td_admin_4"><textarea id="ContentTextHeb" name="ContentTextHeb" dir="rtl" style="HEIGHT:80px;WIDTH:450px;LINE-HEIGHT:120%"
														readonly><%=pContentTextHeb%></textarea></td>
												<td align="right">&nbsp;����� ������ ������</td>
											</tr>
											<tr>
												<td align="right" class="td_admin_4"><textarea id="ContentText" name="ContentText" dir="ltr" style="HEIGHT:80px;WIDTH:450px;LINE-HEIGHT:120%"
														class="clsAlphaNoOnly"><%=pContentText%></textarea></td>
												<td align="right">&nbsp;����� �������</td>
											</tr>
											<tr>
												<td height="20">&nbsp;</td>
											</tr>
											<tr>
												<td align="center"><table border="0" cellpadding="0" cellspacing="0" align="center">
														<tr>
															<td align="center" class="td_admin_4">
																<input type="submit" name="Button1" style="WIDTH:300px" value="���� ����� ������ ����� ����� �� �����"
																	id="Button1" class="but_menu" onClick="return CheckForm()">
															</td>
																													
														</tr>
													</table>
    </form>

  </body>
</html>
