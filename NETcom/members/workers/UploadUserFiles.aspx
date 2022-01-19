<%@ Page Language="vb" AutoEventWireup="false" Codebehind="UploadUserFiles.aspx.vb" Inherits="bizpower_pegasus2018.UploadUserFiles"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
    <title>UploadUserFiles</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../IE4.css" rel="STYLESHEET" type="text/css">
    <script>
    function OpenFile(fname)
    {
    window.open("checkUser.aspx?fname="+ fname, "winp", "width=400,height=100");
    }
    function checkDelete(Id)
	{

		return window.confirm("?האם ברצונך למחוק קובץ");
		
	}
    </script>
  </head>
  <body MS_POSITIONING="GridLayout">

    <form id="Form1" method="post" runat="server">
  <div>
                  <table cellpadding="0" bgcolor="#e6e6e6" cellspacing="0" cellpadding="0" width="100%" border="0" dir="ltr">
 
           <tr>
                    <td style="background-color: #FFD011">
                        <table cellpadding="0" cellspacing="0" width="100%" align="left" style="border: solid 0px #d3d3d3">
                            <tr>
                                <td><a class="button_edit_1" style="width: 95px; font_family: ARIAL; font-size: 14px; font-weight: normal" href="javascript:void(0)" onclick="javascript:window.open('','_self').close();">סגור חלון</a></td>
                                <td>
                                <%if  trim(Request.Cookies("bizpegasus")("UploadUserFiles"))=1  then%>
                                <a href="#"  class="button_edit_1" style="width:200" onclick="window.open('UploadFile.aspx?Userid=<%=UserId%>','winUPD','top=20, left=10, width=1000, height=350, scrollbars=1');">הוספת קובץ</a>
                                <%end if%>
                                </td>
                                <td width="100%" class="page_title" dir="rtl">&nbsp;<font style="color: #000000;"><b></b></font>&nbsp;</td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td align="center" width="100%" valign="top" style="border: 1px solid #808080; border-left: none; border-top: none">

        <table cellpadding="1" cellspacing="2"  width="100%" align="center" style="border:solid 1px #d3d3d3" dir=rtl>
		    	<tr bgcolor="#d8d8d8" style="height:45px">
					         <td class="title_sort" align="center" width=10%>תאריך</td>
                     	     <td class="title_sort" align="center"  width=1%>&nbsp;</td>
                    		 <td class="title_sort" align="center" width =40%>שם קובץ</td>
                     		 <td class="title_sort" align="center" width=30%>תאור</td>
			                 <td class="title_sort" align="center"  width=10% nowrap>שם מעלה הקובץ&nbsp;</td>
			                 <td class="title_sort" align="center" width=5%>מחיקה</td>
						
		        </tr>
 <asp:repeater ID="rptData" runat="server">
           <Itemtemplate>
                  <tr  style="height:30px">
                             <td align="center"><%#DataBinder.Eval(Container.DataItem, "Date", "{0:dd/MM/yyyy }")%> <%#DataBinder.Eval(Container.DataItem, "Date", "{0:HH:MM }")%></td>
                     		 <td align="right">&nbsp;<%#IIf(Len(Container.DataItem("File_Name")) > 0, "<a href=""javascript:OpenFile('"& Container.DataItem("File_Name") &"')""><img src=../../images/MsgAttachment.gif border=0></a>", "")%></td>
							 <td align="right" dir=ltr>&nbsp;<%#IIf(Len(Container.DataItem("File_Name")) > 0, "<a href=""javascript:OpenFile('"& Container.DataItem("File_Name")&"')"">"& Container.DataItem("File_Name") &"</a>", "")%></td>
				 
							 <td align="center"><%#Container.DataItem("File_Description")%><%'#func.breaks(Container.DataItem("File_Description"))%></td>
                             <td align="center"><%#Container.DataItem("Worker_NAME")%></td>
                             <td align=center><%IF trim(Request.Cookies("bizpegasus")("UploadUserFilesDelete"))=1 THEN%><a href="UploadUserFiles.aspx?userId=<%=UserId%>&deleteId=<%#Container.DataItem("Id")%>" onclick="return checkDelete('<%#Container.DataItem("Id")%>')"><img src="../../images/delete_icon.gif" border="0"></a><%end if%></td>
                  </tr>
     </Itemtemplate>
    <Separatortemplate>
        <tr><td colspan="6" style="background-color:#ffffff;height:1px"></td></tr>
    </Separatortemplate>
</asp:repeater>
	
            </table></td></tr></table>
    </div>
    </form>

  </body>
</html>
