<%@ Page Language="vb" AutoEventWireup="false" Codebehind="RoomingListUpload.aspx.vb" Inherits="bizpower_pegasus2018.RoomingListUpload" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>

	<HEAD>
		<title>Rooming List Upload</title>
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
		<input type="hidden" name="DId" id="DId" value="<%=DepartureId%>">
		<table width="100%" align="center" cellspacing="0" cellpadding="2" border="0" >
						<tr>
					<td bgcolor="#e6e6e6" height="5" align="center" width="100%" colspan=2><b>Rooming List Upload – <%=DepatureCode%><%'=func.GetDepatureCode(DepartureId)%> <br></b></td>
				</tr>

		<tr><td width=20% align=center valign=top>

		<table border=0 cellpadding=0 cellspacing=0>
		
		<tr><td height=20>&nbsp;</td></tr>
		<tr><td align="center" style="padding-left:30px" nowrap>
		<div>
		<asp:PlaceHolder  id=checkSup Visible=false runat=server>
		<table  cellpadding=5 cellspacing=5 style="border: #ff0000 solid 1px" align=center>
		<tr><td align=right nowrap><b>!שים לב</b></td></tr>
		<tr><td align=right nowrap>לא קיים מייל מעודכן במערכת לספק </td></tr>
		<tr><td align=center><b><%=texCheck%></b></td></tr>
		<tr><td align=right>להוספת כתובת מייל תיקנית יש לגשת למסך ניהול הספקים</td></tr>
		</table>
		</asp:PlaceHolder>
		</div>
		</td></tr>
		</table>
		</td>
		<td align=center>
			<table width="100%" align="center" cellspacing="0" cellpadding="2" border="0">
			
<tr>
			
					<td align="center" class="td_admin_4" height=10></td></tr>
					<tr>
					<td align="center"><table border=0 >
					<tr>
					<td align="right" colspan=2>
					<asp:PlaceHolder id="divselect" runat=server >
						<select id="fselect" name="fselect"  dir=rtl runat=server  style="width:300px;font-size:9pt;font-family:arial;">
						</select>
						</asp:PlaceHolder>
						<asp:PlaceHolder ID="divfselectError" Runat=server Visible=False><span style="color:#ff0000">הספקים לא  מוגדרים לטיול ספציפי</br>
						לא נתן לעלות קבצים</span>
						</asp:PlaceHolder>
					</td>
				</tr>
			<tr><td align="right" class="td_admin_4">
			
			<input name="dateP" type="text" id="dateP" dir="ltr" runat="server" disabled></td><td align=right>&nbsp;תאריך&nbsp;</td></tr>
			<tr><td align="right" class="td_admin_4"><input name="UserName" type="text" id="UserName" dir="ltr" runat="server" disabled></td><td align=right>&nbsp;שם משתמש&nbsp;</td></tr>
		
				<tr>
					<td align="center" class="td_admin_4" height=10 colspan=2></td></tr>
				  <tr><td align="center" class="td_admin_4"><input name="fileupload1" type="file" id="fileupload1" dir="rtl" runat="server"></td><td align=right>&nbsp;צירוף קובץ&nbsp;</td></tr>
				  <tr><td colspan=2 align="center">jpg,jpeg,png, gif, pdf, docx, doc, xls, xlsx, txt</td></tr>
                <tr>
					<td align="center" class="td_admin_4" height=10 colspan=2></td></tr>
					<tr>
	<td align="right" class="td_admin_4"><input type="checkbox" dir="ltr" name="Guaranteed" ID="Guaranteed" <%if isGuaranteed=True then%> checked <%else%> onClick="javascript:Message(this)" <%end if%>></td>
	<td align="left" class="td_admin_4" ><div class="tooltip">טיול מובטח
 		<asp:Repeater ID=GLog Runat=server Visible=false>
		<HeaderTemplate> <span class="tooltiptext">	<table  cellpadding=1 cellspacing=1 style="border: #ff0000 solid 0px" align=left>
	
	<!--tr><td align=left nowrap class="td_admin_4">User Name</td>
		<td align=left nowrap class="td_admin_4">Date </td>
		<td align=center class="td_admin_4">&nbsp;</td>
		</tr--></HeaderTemplate>
		<ItemTemplate><tr><td align=left ><%#Container.DataItem("UserName")%> </td>
		<td align=left><%#Container.DataItem("Date")%> </td>
		<td align=left><%#IIF(Container.DataItem("Guaranteed")=True,"Guaranteed","not Guaranteed")%> </td>
		</tr></ItemTemplate>
		<SeparatorTemplate><TR><td colspan=3 style="height:1px" width=100% bgcolor=#ffffff></td></TR></SeparatorTemplate>
		<FooterTemplate></table></span>
</div>&nbsp;</FooterTemplate>
		</asp:Repeater>
		
</td>
</tr>
				<tr>
					<td align="center" class="td_admin_4" colspan=2>
					<input type="button" name="Button1" style="width:100px" value="הוסף קובץ" id="Button1"  class="but_menu" onClick="return CheckForm()" />
					</td>
				</tr>
			</table>
			</td></tr></table>
			</td></tr></table>
		</form>
	<%if false then%>	<div id=resultF name="resultF" >
			<table width="100%" align="center" cellspacing="0" cellpadding="2" border="0">
				<tr>
					<td bgcolor="#e6e6e6" height="5" align="center" width="100%"><b>Rooming List Upload – <%=DepatureCode%><%''=func.GetDepatureCode(DepartureId)%> <br></b></td>
				</tr>
				<tr><td height=20></td></tr>
				<tr><TD id=resultSup class="td_admin_4" align=center>הקובץ הועלה בהצלחה לספקים הבאים: <%=texCheck%></TD></tr>
				<tr><td height=20></td></tr>
				
				<tr><td align =center class="td_admin_4"><a href=# class="but_menu" style="background-color:#e1e1e1;height:30;text-decoration:none;color:#000000;padding:3px"  onclick="javascript:self.close();window.opener.location.reload(true);">סגור דף</a></td></tr>
				</table>
		</div><%end if%>
	</body>
</HTML>
