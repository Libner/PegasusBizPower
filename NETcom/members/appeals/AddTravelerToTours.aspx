<%@ Page Language="vb" AutoEventWireup="false" Codebehind="AddTravelerToTours.aspx.vb" Inherits="bizpower_pegasus2018.AddTravelerToTours"%>
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
	<script>
	function checktransf()
	{
	//alert("checktransf")
	  var arrid=document.getElementsByName("chkTour")
        var par =''
        for (i=0;i<arrid.length;i++)
        {
          //alert(arrid[i].id+' '+arrid[i].checked)
            if (arrid[i].checked)
            {
                // alert(arrid[i].id)
                //  alert(arrid[i].id)
                if (par=='')
                {par=arrid[i].id.substring(7)
                }

                else{
                    par=par+','+ arrid[i].id.substring(7)
                }
                //alert(arrid[i].id.substring(4))
            }

        }
        
        if (par!='')
        {
          //  window.open ("ScreenTable.aspx?"+"par="+ par,"ScreenTable","scrollbars=yes,menubar=yes, toolbar=yes, width=900,height=800,left=10,top=10");
              $.ajax
({
    type: 'post',
    url: 'AddTravelersrecord.aspx',
    data: {
        edit_row: 'update_row',
        traveler_id:par,
        appdocket:<%=appDocket%>,
        appid:<%=appid%>,
        depId:<%=departureId%>
          },
    async:false,
   
    success: function (response) {
    alert(response)
         var str = response;
         window.close()
opener.location.reload();		 }
       
});
        }
        else
        {
            alert("אנ לסמן רשומות ")
            return false;
        }
      
    }
        

	</script>
  </head>
  <body MS_POSITIONING="GridLayout">

    <form id="Form1" method="post" runat="server">
    <input type="hidden" id="appDocket" name="appDocket" value="<%=appDocket%>">
    <input type="hidden" id="appid" name="appid" value="<%=appid%>">
      <div>
    <table cellpadding=0 cellspacing=0 width=100% border=0>
    <tr>
    <td width=100% valign=top>
 <table cellpadding="1" cellspacing="1" align="center" style="width:80%;border:solid 0px #d3d3d3">
	 <tr>
         <td height="30" align="center" colspan="9"><span style="COLOR: #6F6DA6;font-size:14pt"></span></td>
							</tr>
     
           <tr>
             <td style="background-color:#bcbbd5" align="center"><input name="sIdNumber" class="searchList" id="sIdNumber" style="WIDTH:80px;background-color:#ffffff" onkeypress="return getNumbersCharts(this)" type="text" runat="server"></td>
         	<td class="title_sort" align="center" style="background-color:#bcbbd5"><input name="sLastName" class="searchList" id="sLastName" style="WIDTH:120px;background-color:#ffffff" onkeypress="return getNumbersCharts(this)" type="text" runat="server"></td>
			<td class="title_sort" align="center" style="background-color:#bcbbd5"><input name="sFirstName" class="searchList" id="sFirstName" style="WIDTH:120px;background-color:#ffffff" onkeypress="return getNumbersCharts(this)" type="text" runat="server"></td>
            <td class="title_sort" align="center"  style="background-color:#bcbbd5;vertical-align:top">
                <table border="0" cellpading="0" cellspacing="0">
                    <tr><td class="title_sort" align="center" style="background-color:#bcbbd5;vertical-align:middle"><a href="" onclick="cal1xx.select(document.getElementById('sBirthdayDate'),'AsBirthdayDate','dd/MM/yyyy'); return false;" id ="AsBirthdayDate"><img src="../../images/calendar.gif" border="0" style="padding-top:5px"></a></td>
                        <td class="title_sort" align="center" style="background-color:#bcbbd5;vertical-align:middle" ><input runat="server" type="text" id="sBirthdayDate" class="searchList" style="WIDTH:80px;background-color:#ffffff" onkeypress="return getNumbersCharts(this)" name="sBirthdayDate"></td>
                    </tr>

                </table>
                </td>
        	<td class="title_sort" align="center" style="background-color:#bcbbd5;"><input name="sPasportNumber" class="searchList" id="sPasportNumber" style="WIDTH:120px;background-color:#ffffff" onkeypress="return getNumbersCharts(this)" type="text" runat="server"></td>
        	<td class="title_sort" align="center" style="background-color:#bcbbd5"><input name="sMobileNumber" class="searchList" id="sMobileNumber" style="WIDTH:120px;background-color:#ffffff" onkeypress="return getNumbersCharts(this)" type="text" runat="server"></td>
	<td  class="title_sort" align="center"><asp:LinkButton Runat="server" ID="btnSearch" CssClass="button_small1">חפש</asp:LinkButton></td>
         
     </tr>
     <tr style="height:30px">
    	<td class="title_sort" align="center">ID number
                                    <a href="javascript:FuncSort('<%=qrystring%>','sort_1','ASC')"><img src="../../../images/arrow_top<%If Request.QueryString("sort_1") = "ASC" Then%>_act<%End If%>.gif"  title="למיין לפי ת.ז" border=0></a>
                                    <a href="javascript:FuncSort('<%=qrystring%>','sort_1','DESC')"><img src="../../../images/arrow_bot<%If Request.QueryString("sort_1") = "DESC" Then%>_act<%End If%>.gif"  title="למיין לפי ת.ז" border=0></a>
		</td>
		<td class="title_sort" align="center">Last Name (English)
                                    <a href="javascript:FuncSort('<%=qrystring%>','sort_2','ASC')"><img src="../../../images/arrow_top<%If Request.QueryString("sort_2") = "ASC" Then%>_act<%End If%>.gif"  title="למיין לפי סטטוס" border=0></a>
                                    <a href="javascript:FuncSort('<%=qrystring%>','sort_2','DESC')"><img src="../../../images/arrow_bot<%If Request.QueryString("sort_2") = "DESC" Then%>_act<%End If%>.gif"  title="למיין לפי שם משפחה" border=0></a>
        </td>
		<td class="title_sort" align="center">First Name (English)
                                    <a href="javascript:FuncSort('<%=qrystring%>','sort_3','ASC')"><img src="../../../images/arrow_top<%If Request.QueryString("sort_3") = "ASC" Then%>_act<%End If%>.gif"  title="למיין לפי שם" border=0></a>
                                    <a href="javascript:FuncSort('<%=qrystring%>','sort_3','DESC')"><img src="../../../images/arrow_bot<%If Request.QueryString("sort_3") = "DESC" Then%>_act<%End If%>.gif"  title="למיין לפי שפ" border=0></a>
        </td>
        <td class="title_sort" align="center">Date of Birthday (DD/MM/YY)
                                    <a href="javascript:FuncSort('<%=qrystring%>','sort_4','ASC')"><img src="../../../images/arrow_top<%If Request.QueryString("sort_4") = "ASC" Then%>_act<%End If%>.gif"  title="למיין לפי תאריך לידה" border=0></a>
                                    <a href="javascript:FuncSort('<%=qrystring%>','sort_4','DESC')"><img src="../../../images/arrow_bot<%If Request.QueryString("sort_4") = "DESC" Then%>_act<%End If%>.gif"  title="למיין לפי תאריך לידה" border=0></a>
        </td>
        <td class="title_sort" align="center">Pasport Number
                                    <a href="javascript:FuncSort('<%=qrystring%>','sort_5','ASC')"><img src="../../../images/arrow_top<%If Request.QueryString("sort_5") = "ASC" Then%>_act<%End If%>.gif"  title="למיין לפי מספר דרכון" border=0></a>
                                    <a href="javascript:FuncSort('<%=qrystring%>','sort_5','DESC')"><img src="../../../images/arrow_bot<%If Request.QueryString("sort_5") = "DESC" Then%>_act<%End If%>.gif"  title="למיין לפי מספר דרכון" border=0></a>
        </td>
        <td class="title_sort" align="center">Mobile Number
                                    <a href="javascript:FuncSort('<%=qrystring%>','sort_6','ASC')"><img src="../../../images/arrow_top<%If Request.QueryString("sort_6") = "ASC" Then%>_act<%End If%>.gif"  title="למיין לפי טלפון" border=0></a>
                                    <a href="javascript:FuncSort('<%=qrystring%>','sort_6','DESC')"><img src="../../../images/arrow_bot<%If Request.QueryString("sort_6") = "DESC" Then%>_act<%End If%>.gif"  title="למיין לפי טלפון" border=0></a>
        </td> <td  class="title_sort" align="center">&nbsp;</td>
   	
    </tr>
 
      <asp:Repeater ID="rptData" runat="server">
     <ItemTemplate>
        <tr style="background-color: rgb(201, 201, 201);height:30px;" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#C9C9C9';" style="background-color: rgb(201, 201, 201);">
          
            <td align="center"><%#Container.DataItem("IdNumber")%></td>
            <td align="center"><%#Container.DataItem("LastName")%></td>
            <td align="center"><%#Container.DataItem("FirstName")%></td>
            <td align="center"><%#DataBinder.Eval(Container.DataItem,"BirthDate", "{0:d/M/yyyy}")%></td>
            <td align="center"><%#Container.DataItem("PassportNum")%></td>
            <td align="center"><%#Container.DataItem("Phone")%></td>
            <td  align="center">
            <%'Modified by Mila 27/04/20 - check only traveler is not joinde to departure (not only to docket)%>
            <%'#IIF(func.dbNullFix(Container.dataItem("Docket_pfileNum"))=appDocket,"","<input type=""checkbox"" ID=""chkTour" & Container.DataItem("Traveler_Id")&""" NAME=""chkTour""  title=""הוסף מטייל לטיול"">")%>
            <%#IIF(func.fixnumeric(Container.dataItem("isJoinedToDeparture"))=1,"","<input type=""checkbox"" ID=""chkTour" & Container.DataItem("Traveler_Id")&""" NAME=""chkTour""  title=""הוסף מטייל לטיול"">")%>
                                </td>
       </tr>
     </ItemTemplate>
     </asp:Repeater>
       <tr><TD colspan=7 align=center height=20>&nbsp;</td></tr>
     <tr><TD colspan=7 align=center>
     <table border=0 cellpadding=1 cellspacing=1 align=center>
     <tr><td nowrap width=100 > <a class="button_edit_1" onclick =window.close()>&nbsp;&nbsp;&nbsp;&nbsp;בטל&nbsp;&nbsp;&nbsp;&nbsp;</a></td> 
     <td>&nbsp;&nbsp;&nbsp;</td>
     <td nowrap>  <a class="button_edit_1" onclick="checktransf()" href="#" nowrap="">&nbsp;&nbsp;קישור לטיול&nbsp;&nbsp;</a></td></tr>
     </table>
   </td></tr>
     </table></td></tr></table>
    </div>
      <SCRIPT type="text/javascript">
            <!--
    var cal1xx = new CalendarPopup('CalendarDiv');
    cal1xx.setYearSelectStartOffset(120);
    cal1xx.setYearSelectEndOffset(2);
    //cal1xx.setYearSelectStart(1910); //Added by Mila
    cal1xx.showNavigationDropdowns();
    cal1xx.offsetX = -50;
    cal1xx.offsetY = -40;

    //-->
    </SCRIPT>
	<DIV ID='CalendarDiv' STYLE='POSITION:absolute;VISIBILITY:hidden;BACKGROUND-COLOR:white;layer-background-color:white'></DIV>

    </form>

  </body>
</html>
