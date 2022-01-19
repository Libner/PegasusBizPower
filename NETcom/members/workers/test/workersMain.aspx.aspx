<%@ Page Language="vb" AutoEventWireup="false" Codebehind="workers.aspx.vb" Inherits="bizpower_pegasus.workers"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
  <head>
    <title>workers</title>
 	<LINK href="../../IE4.css" type="text/css" rel="STYLESHEET">
	     
<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.8.2/jquery.min.js"></script> 
<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.9.1/jquery-ui.min.js"></script> 
<script type="text/javascript" src="gridviewScroll.min.js"></script> 


<script type="text/javascript">
    $(document).ready(function () {
        gridviewScroll();
    });

    function gridviewScroll() {
  // alert($('#frame').contents().find('body').innerWidth())
        screenW = $("body").innerWidth();
//alert(screenW)
        $('#Table1').gridviewScroll({
            width: screenW-6,
            height: 570,
            freezesize: 2
        });
    } 
</script> 
<style>
.GridviewScrollHeader TH, .GridviewScrollHeader TD 
{ 
    padding: 5px; 
    font-weight: bold; 
    white-space: nowrap; 
    border-right: 1px solid #AAAAAA; 
    border-bottom: 1px solid #AAAAAA; 
    background-color: #EFEFEF; 
    text-align: left; 
    vertical-align: bottom; 
} 
.GridviewScrollItem TD 
{ 
    padding: 5px; 
    white-space: nowrap; 
    border-right: 1px solid #AAAAAA; 
    border-bottom: 1px solid #AAAAAA; 
    background-color: #FFFFFF; 
} 
.GridviewScrollPager  
{ 
    border-top: 1px solid #AAAAAA; 
    background-color: #FFFFFF; 
} 
.GridviewScrollPager TD 
{ 
    padding-top: 3px; 
    font-size: 14px; 
    padding-left: 5px; 
    padding-right: 5px; 
} 
.GridviewScrollPager A 
{ 
    color: #666666; 
}
.GridviewScrollPager SPAN

{

    font-size: 16px;

    font-weight: bold;

}
</style>


 
  </head>
	<body style="margin:0px">
    <form id="Form1" method="post" runat="server">
<table cellspacing="0" id="Table1" style="width:100%;border-collapse:collapse;"> 

				<asp:Repeater id="rptWorkers" runat="server" >
	<HeaderTemplate>
    <tr class="GridviewScrollHeader"> 
        <th scope="col">שם עובד</th> 
        <th scope="col">סוג עובד</th> 
        <th scope="col">יעד הזמנות מינימלי</th> 
        <th scope="col">שווי כספי של הזמנה</th> 
        <th scope="col">דרגת חשיבות ידיעת המשוב</th> 
        <th scope="col">StandardCost</th> 
        <th scope="col">ListPrice</th> 
        <th scope="col">Weight</th> 
        <th scope="col">SellStartDate</th> 
        <th scope="col"><%If 0 Then%>מחיקה<%End If%></th>
	<th scope="col" align="center" class="title_sort">עדכון</th>

    </tr> 
    </HeaderTemplate>	
    <ItemTemplate> 
     <tr class="GridviewScrollItem"> 
      <td style="background-color: #EFEFEF;"><%#Container.DataItem("user_name")%></td> 
        <td style="background-color: #EFEFEF;">HL Mountain Frame - Black, 38</td> 

         <td>FR-M94B-38</td> 
        <td>500</td> 
        <td>375</td> 
        <td>739.0410</td> 
        <td>1349.6000</td> 
        <td>2.68</td> 
        <td>7/1/2005 12:00:00 AM</td> 
     <td dir="<%=dir_obj_var%>" align="center" class="card">
		<%If 0 Then%>
		<%'If count_tasks = 0 And count_appeals = 0 Then%>
		<a href="default.asp?delUserID=<%#Container.DataItem("USER_ID")%>" ONCLICK="return CheckDel()"><img src="../../images/delete_icon.gif" border="0" alt="מחיקה"></a>
		<%'Else%>
		<%If trim(lang_id) = "1" Then
			str_alert = "שים לב, קיימת מידע במערכת עבור עובד זה\n\n" & Space(3) & "לפי כך לא ניתן למחוק את העובד ממערכת\n\n" & Space(4) & "אלא להעביר את העובד לסטטוס לא פעיל"
		Else
			str_alert = "Pay attention,\n\n you can\'t delete this employee \n\n however you can deactivate him"
		End If%>		
		<input type=image src="../../images/delete_icon.gif" border=0 Onclick="window.alert('<%=str_alert%>');return false;">
		<%'End If%>
		<%End If%>		
		</td>
		<td dir="<%=dir_obj_var%>" align="center" class="card"><a href="addWorker.asp?USER_ID=<%#Container.DataItem("USER_ID")%>" target=_self><img src="../../images/edit_icon.gif" border="0"></a></td>
    </tr>
     </ItemTemplate>	
  
        </asp:Repeater>
</table> 
    </form>

  </body>
</html>
