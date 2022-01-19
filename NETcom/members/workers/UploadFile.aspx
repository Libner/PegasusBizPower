<%@ Page Language="vb" AutoEventWireup="false" Codebehind="UploadFile.aspx.vb" Inherits="bizpower_pegasus2018.UploadFile"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>

	<HEAD>
		<title>Upload file</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
		<style>
.tooltip {
    position: relative;
    display: inline-block;
    border-bottom: 1px dotted black;
}

.tooltip .tooltiptext {
    visibility: hidden;
    width: 280px;
    background-color: #e1e1e1;
    color: #000;
    text-align: center;
    border-radius: 6px;
    padding: 5px 0;
    position: absolute;
    z-index: 1;
    top: 125%;
    left: 20%;
    margin-left: -60px;
    opacity: 0;
    transition: opacity 0.3s;
}

.tooltip .tooltiptext::after {
    content: "";
    position: absolute;
    top: 100%;
    left: 20%;
    margin-left: -5px;
    border-width: 5px;
    border-style: solid;
    border-color: #555 transparent transparent transparent;
}

.tooltip:hover .tooltiptext {
    visibility: visible;
    opacity: 1;
}
</style>
		<script>
		function Message(e)
		{
	
	if (confirm('הפיכת הטיול למובטח תגרום לספק להתחייב על כל שירותי הטיול, האם ברצונך להמשיך בתהליך ?')) {
       document.getElementById("Guaranteed").checked=true;
} else {
    document.getElementById("Guaranteed").checked=false;
}
	
		}
		
	
	    function CheckForm()
        {
        
              if (document.getElementById("fileupload1").value == '') 
            {
                  alert("אנא בחר  קובץ");
                  document.getElementById("fileupload1").focus()
                return false;
            }
           document.getElementById("Form1").submit();
            return true;

        }
        </script>
	</HEAD>
	<body style="MARGIN:0px" onload="self.focus();">
		<form id="Form1" method="post" runat="server">
		<input type="hidden" name="Userid" id="Userid" value="<%=Userid%>">
		<table width="100%" align="center" cellspacing="0" cellpadding="2" border="0" >
						<tr>
					<td bgcolor="#e6e6e6" height="5" align="center" width="100%" colspan=2><b>העלת קובץ למשתמש:  <%=func.GetSelectUserName(UserId)%><br></b></td>
				</tr>

		<tr><td width=20% align=center valign=top>
    	</td>
		<td align=center>
			<table width="100%" align="center" cellspacing="0" cellpadding="2" border="0">
			
<tr>
			
					<td align="center" class="td_admin_4" height=10></td></tr>
					<tr>
					<td align="center"><table border=0 >
				
			<tr><td align="right" class="td_admin_4">
			
			<input name="dateP" type="text" id="dateP" dir="ltr" runat="server" disabled></td><td align=right>&nbsp;תאריך&nbsp;</td></tr>
			<tr><td align="right" class="td_admin_4"><input name="UserName" type="text" id="UserName" dir="rtl" runat="server" disabled ></td>
			<td align=right>&nbsp;שם משתמש&nbsp;</td></tr>
		<tr>
					<td align="center" class="td_admin_4" height=10 colspan=2></td></tr>
					<tr><td   align="right" class="td_admin_4"><textarea id=fileDescr runat=server style="width:250" dir=rtl></textarea></td>
					<td align=right>&nbsp;תאור&nbsp;</td>
					</tr>
		
				<tr>
					<td align="center" class="td_admin_4" height=10 colspan=2></td></tr>
				  <tr><td align="center" class="td_admin_4"><input name="fileupload1" type="file" id="fileupload1" dir="rtl" runat="server"></td>
				  <td align=right>&nbsp;צירוף קובץ&nbsp;</td></tr>
				  <tr><td colspan=2 align="center">jpg,jpeg,png, gif, pdf, docx, doc, xls, xlsx, txt, msg</td></tr>
                <tr>
					<td align="center" class="td_admin_4" height=10 colspan=2></td></tr>
				
				<tr>
					<td align="center" class="td_admin_4" colspan=2>
					<input type="button" name="Button1" style="width:100px" value="הוסף קובץ" id="Button1"  class="but_menu" onClick="return CheckForm()" />
					</td>
				</tr>
			</table>
			</td></tr></table>
			</td></tr></table>
		</form>
	
	</body>
</HTML>
