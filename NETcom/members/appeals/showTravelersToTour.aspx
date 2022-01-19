<%@ Page Language="vb" AutoEventWireup="false" Codebehind="showTravelersToTour.aspx.vb" Inherits="bizpower_pegasus2018.showTravelersToTour"%>
<%@ Register TagPrefix="DS" TagName="metaInc" Src="../../title_meta_inc.ascx" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
    <title>AddTravelerToTours</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
	<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
    <script src="../../include/CalendarPopup.js" charset="windows-1255"></script>
	<DS:metaInc id="metaInc" runat="server"></DS:metaInc>    

  </head>
  <body MS_POSITIONING="GridLayout">

    <form id="Form1" method="post" runat="server">
    <input type="hidden" id="appDocket" name="appDocket" value="<%=appDocket%>">
    <input type="hidden" id="appid" name="appid" value="<%=appid%>">
      <div>
    <table cellpadding=0 cellspacing=0 width=100% border=0>
    <tr>
					<td class="title_form" style="font-size:14pt;line-height:24px" width="100%" align="center" dir="rtl">
					מטיילים בקבוצה <%=departure_Code%>
					</td>
				</tr>
			
     <TD  align=center height=20>&nbsp;</td></tr>
    <tr>
    <td width=100% valign=top>
 <table cellpadding="1" cellspacing="1" align="center" style="width:80%;border:solid 0px #d3d3d3">
	
     <tr style="height:30px">
	<td class="title_sort" align="center">&nbsp;</td>
   		<td class="title_sort" align="center">Docket number
         <!--                           <a href="javascript:FuncSort('<%=qrystring%>','sort_7','ASC')"><img src="../images/arrow_top<%If Request.QueryString("sort_7") = "ASC" Then%>_act<%End If%>.gif"  title="למיין לפי מספר דוקט" border=0></a>
                                    <a href="javascript:FuncSort('<%=qrystring%>','sort_7','DESC')"><img src="../images/arrow_bot<%If Request.QueryString("sort_7") = "DESC" Then%>_act<%End If%>.gif"  title="למיין לפי מספר דוקט" border=0></a>
		-->
   		</td>
		<td class="title_sort" align="center">Room join number
         <!--                           <a href="javascript:FuncSort('<%=qrystring%>','sort_8','ASC')"><img src="../images/arrow_top<%If Request.QueryString("sort_8") = "ASC" Then%>_act<%End If%>.gif"  title="למיין לפי מספר חדר" border=0></a>
                                    <a href="javascript:FuncSort('<%=qrystring%>','sort_8','DESC')"><img src="../images/arrow_bot<%If Request.QueryString("sort_8") = "DESC" Then%>_act<%End If%>.gif"  title="למיין לפי מספר חדר" border=0></a>
		-->
   		</td>
		<td class="title_sort" align="center">Room type
         <!--                           <a href="javascript:FuncSort('<%=qrystring%>','sort_9','ASC')"><img src="../images/arrow_top<%If Request.QueryString("sort_9") = "ASC" Then%>_act<%End If%>.gif"  title="למיין לפי מספר חדר" border=0></a>
                                    <a href="javascript:FuncSort('<%=qrystring%>','sort_9','DESC')"><img src="../images/arrow_bot<%If Request.QueryString("sort_9") = "DESC" Then%>_act<%End If%>.gif"  title="למיין לפי מספר חדר" border=0></a>
		-->
   		</td>
		<td class="title_sort" align="center">ID number
        <!--                            <a href="javascript:FuncSort('<%=qrystring%>','sort_1','ASC')"><img src="../images/arrow_top<%If Request.QueryString("sort_1") = "ASC" Then%>_act<%End If%>.gif"  title="למיין לפי ת.ז" border=0></a>
                                    <a href="javascript:FuncSort('<%=qrystring%>','sort_1','DESC')"><img src="../images/arrow_bot<%If Request.QueryString("sort_1") = "DESC" Then%>_act<%End If%>.gif"  title="למיין לפי ת.ז" border=0></a>
		-->
   		</td>
		<td class="title_sort" align="center">Last Name (English)
        <!--                            <a href="javascript:FuncSort('<%=qrystring%>','sort_2','ASC')"><img src="../images/arrow_top<%If Request.QueryString("sort_2") = "ASC" Then%>_act<%End If%>.gif"  title="למיין לפי סטטוס" border=0></a>
                                    <a href="javascript:FuncSort('<%=qrystring%>','sort_2','DESC')"><img src="../images/arrow_bot<%If Request.QueryString("sort_2") = "DESC" Then%>_act<%End If%>.gif"  title="למיין לפי שם משפחה" border=0></a>
        -->
   		</td>
		<td class="title_sort" align="center">First Name (English)
        <!--                            <a href="javascript:FuncSort('<%=qrystring%>','sort_3','ASC')"><img src="../images/arrow_top<%If Request.QueryString("sort_3") = "ASC" Then%>_act<%End If%>.gif"  title="למיין לפי שם" border=0></a>
                                    <a href="javascript:FuncSort('<%=qrystring%>','sort_3','DESC')"><img src="../images/arrow_bot<%If Request.QueryString("sort_3") = "DESC" Then%>_act<%End If%>.gif"  title="למיין לפי שפ" border=0></a>
        -->
   		</td>
        <td class="title_sort" align="center">Date of Birthday (DD/MM/YY)
        <!--                            <a href="javascript:FuncSort('<%=qrystring%>','sort_4','ASC')"><img src="../images/arrow_top<%If Request.QueryString("sort_4") = "ASC" Then%>_act<%End If%>.gif"  title="למיין לפי תאריך לידה" border=0></a>
                                    <a href="javascript:FuncSort('<%=qrystring%>','sort_4','DESC')"><img src="../images/arrow_bot<%If Request.QueryString("sort_4") = "DESC" Then%>_act<%End If%>.gif"  title="למיין לפי תאריך לידה" border=0></a>
        -->
   		</td>
        <td class="title_sort" align="center">Pasport Number
        <!--                            <a href="javascript:FuncSort('<%=qrystring%>','sort_5','ASC')"><img src="../images/arrow_top<%If Request.QueryString("sort_5") = "ASC" Then%>_act<%End If%>.gif"  title="למיין לפי מספר דרכון" border=0></a>
                                    <a href="javascript:FuncSort('<%=qrystring%>','sort_5','DESC')"><img src="../images/arrow_bot<%If Request.QueryString("sort_5") = "DESC" Then%>_act<%End If%>.gif"  title="למיין לפי מספר דרכון" border=0></a>
        -->
   		</td>
        <td class="title_sort" align="center">Mobile Number
         <!--                           <a href="javascript:FuncSort('<%=qrystring%>','sort_6','ASC')"><img src="../images/arrow_top<%If Request.QueryString("sort_6") = "ASC" Then%>_act<%End If%>.gif"  title="למיין לפי טלפון" border=0></a>
                                    <a href="javascript:FuncSort('<%=qrystring%>','sort_6','DESC')"><img src="../images/arrow_bot<%If Request.QueryString("sort_6") = "DESC" Then%>_act<%End If%>.gif"  title="למיין לפי טלפון" border=0></a>
        -->
   		</td>
   
    </tr>

     <asp:Repeater ID="rptData" runat="server">
     <ItemTemplate>
         <asp:PlaceHolder ID="phSeparateRoomsLine" runat="server" Visible="false">
             <tr><td colspan="11" bgcolor="#ffffff" height="2"></td></tr>
         </asp:PlaceHolder>
        <tr style="background-color: rgb(201, 201, 201);height:30px;" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#C9C9C9';" style="background-color: rgb(201, 201, 201);">
            <td  align="center">&nbsp;<%#Container.ItemIndex+1 %>&nbsp;</td>
              <td align="center"><%#Container.DataItem("Tdocket")%></td>
              <td align="center"><%#Container.DataItem("joinRoomNum")%></td>
              <td align="center"><asp:Literal ID="ltRoomTypeDesc" runat="server"></asp:Literal></td>
            <td align="center"><%#Container.DataItem("IdNumber")%></td>
            <td align="center"><%#Container.DataItem("LastName")%></td>
            <td align="center"><%#Container.DataItem("FirstName")%></td>
            <td align="center"><%#DataBinder.Eval(Container.DataItem,"BirthDate", "{0:d/M/yyyy}")%></td>
            <td align="center"><%#Container.DataItem("PassportNum")%></td>
            <td align="center"><%#Container.DataItem("Phone")%></td>
       </tr>
     </ItemTemplate>
     </asp:Repeater>
       <tr>
       
     </table></td></tr>
     <TD  align=center height=20>&nbsp;</td></tr>
     <tr><TD  align=center>
     <table border=0 cellpadding=1 cellspacing=1 align=center>
     <tr><td nowrap width=100 > <a class="button_edit_1" onclick =window.close()>&nbsp;&nbsp;&nbsp;&nbsp;סגור&nbsp;&nbsp;&nbsp;&nbsp;</a></td> 
        </table>
   </td></tr></table>
    </div>
     
    </form>

  </body>
</html>
