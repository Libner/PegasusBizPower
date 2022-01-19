<%@ Page Language="vb" AutoEventWireup="false" Codebehind="workersTmp.aspx.vb" Inherits="bizpower_pegasus.workersTmp" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
    <title>workersTmp</title>
 
  <meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
  <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.10.15/css/jquery.dataTables.min.css">
<link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/fixedcolumns/3.2.2/css/fixedColumns.dataTables.min.css">
<script type="text/javascript" src="//code.jquery.com/jquery-1.12.4.js"></script>
<script type="text/javascript" src="jquery.dataTables.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/fixedcolumns/3.2.2/js/dataTables.fixedColumns.min.js"></script>
<script>
function OpenEmail(Id)
{
alert("OpenEmail")
}
function addClass(el, className) {

el.className += " " + className
}

function removeClass(el, className) {
   
  if (el.classList)
    el.classList.remove(className)
  else if (hasClass(el, className)) {
    var reg = new RegExp('(\\s|^)' + className + '(\\s|$)')
  
  
  }
    el.className=el.className.replace(className, ' ')
    
}
</script>
<script>

$(document).ready(function() {

 //screenW = $("body").innerWidth();
 //alert(screenW)

 //alert(document.getElementById("dl").className);



    var table = $('#example').DataTable( {
        scrollY:        "500px",
        scrollX:        true,
        scrollCollapse: true,
        paging:         false,
        searching: false,
   "order": [],
    "columnDefs": [ {
      "targets"  : 'no-sort',
      "orderable": false,
    }],
        fixedColumns:   {
            leftColumns: 2,
            rightColumns: 1
        }
     } );

// alert(document.getElementById("dl").className);
 //$("#dl").removeClass("sorting").addClass("main");

} );

</script>
<style>
BODY
{
    PADDING-RIGHT: 0px;  PADDING-LEFT: 0px; PADDING-BOTTOM: 0px;  MARGIN: 0pt 10px; PADDING-TOP: 0px;
    BACKGROUND-COLOR: #ffffff;  FONT-FAMILY: Arial; 
   
}
TD
{
    FONT-SIZE: 12px;
    text-align:center;
    FONT-FAMILY: Arial
}
/* Ensure that the demo table scrolls */
    th, td { white-space: nowrap; }
    div.dataTables_wrapper {
        width: 500px;/*70%;*/
        margin: 0 auto;
    }
</style>
  </head>
  <body >

    <form id="Form1" method="post" runat="server">
<table id="example" class="stripe row-border order-column" cellspacing="0" width="100%">
     				<asp:Repeater id="rptWorkers" runat="server" >
	<HeaderTemplate>

        <thead>
            <tr >
                <th nowrap id="dl" class="no-sort" >מחיקה</th>
                <th  nowrap class="no-sort">עדכון</th>
                 <th nowrap>Position------</th>
                <th nowrap>Office------</th>
                 <th nowrap>Position------</th>
                <th nowrap>Office------</th>
                <th nowrap>Position------</th>
                <th nowrap>Office------</th>
                <th nowrap>דרגת חשיבות ידיעת המשוב</th>
                <th>שווי כספי <BR> של הזמנה</th>
                <th>יעד הזמנות <BR>מינימלי</th>
                 <th nowrap>מחלקה</th>
                 <th nowrap>מייל</th>
                <th nowrap>סוג עובד</th>
                <th nowrap>שם עובד</th>
            </tr>
        </thead>
        
        <tbody>
        </HeaderTemplate>
          <ItemTemplate> 
            <tr>
                <td>Tiger</td>
                <td>Nixon</td>
                <td>System Architect</td>
                <td>Edinburgh</td>
                 <td>System Architect</td>
                <td>Edinburgh</td>
                 <td>System Architect</td>
                <td>Edinburgh</td>
                <td><%#func.dbNullFix(Container.DataItem("ImportanceId"))%></td>
                <td><%#Container.DataItem("Order_Price")%></td>
                <td><%#Container.DataItem("Month_Min_Order")%></td>
                 <td><%#Container.DataItem("email")%></td>
                <td id="Td_<%#Container.DataItem("User_Id")%>" onclick="javascript:OpenEmail()"><input type=text dir="ltr" name="email" value="<%#Container.DataItem("email")%>"  style="width:200;display:none" maxlength=100 style="font-family:arial" ><%#Container.DataItem("email")%></td>
                <td><%#Container.DataItem("job_name")%></td>
                <td><%#Container.DataItem("user_name")%></td>
            </tr>
           </ItemTemplate>
          <FooterTemplate> </tbody></FooterTemplate> 
          </asp:Repeater>
       
    </table>
</form>

</body>

</html>
