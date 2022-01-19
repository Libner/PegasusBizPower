<%@ Page Language="vb" AutoEventWireup="false" Codebehind="workersTmp.aspx.vb" Inherits="bizpower_pegasus2018.workersTmp" %>
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
    <link rel="stylesheet" type="text/css" href="jquery.dataTables.min.css"><!--https://cdn.datatables.net/1.10.15/css/jquery.dataTables.min.css-->
<link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/fixedcolumns/3.2.2/css/fixedColumns.dataTables.min.css">
<script type="text/javascript" src="//code.jquery.com/jquery-1.12.4.js"></script>
<script type="text/javascript" src="jquery.dataTables.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/fixedcolumns/3.2.2/js/dataTables.fixedColumns.min.js"></script>
<script>
function changeInput(n,value)
{
//update value-------
	var  formData = "name="+n+"&Uid="+value+"&value="+document.getElementById(n+"_"+ value).value; 
//alert (formData);
 
$.ajax({
    url : "UpdateData.aspx",
    type: "POST",
    data : formData,
    success: function(data, textStatus, jqXHR)
    {
   //alert(data)
   // var ResData new Array();
   // var ResData=data.split("_");
   //data - response from server
    },
    error: function (jqXHR, textStatus, errorThrown)
    {
// alert(textStatus)
    }
});
	

////
 document.getElementById("dv"+n+"_"+value).style.display="none";
		document.getElementById("dvT"+n+"_"+value).style.display="block";
 document.getElementById("dvT"+n+"_"+value).innerHTML=document.getElementById(n+"_"+ value).value;

}
function ChangeSelect(id)
		{
		var  formData = "id="+id+"&selV="+document.getElementById("sel_"+ id).value; 
	//	alert(formData)
		
$.ajax({
    url : "UpdateSelect.aspx",
    type: "POST",
    data : formData,
    success: function(data, textStatus, jqXHR)
    {
//   alert(data)
//alert(textStatus)
   // var ResData new Array();
  //  var ResData=data.split("_");
    
 //alert(ResData[0]);
 document.getElementById("typeName_"+id).innerHTML=data; //ResData[0];
 document.getElementById("dvDep_"+id).style.display="none";
		document.getElementById("dvTDep_"+id).style.display="block";
        //data - response from server
    },
    error: function (jqXHR, textStatus, errorThrown)
    {
 
    }
});
		
		

		//alert(formData)
		
		//document.getElementById("dvTDep_"+id).style.display="block";
		//alert(document.getElementById("dvTDep_1233").value)
		//document.getElementById("dvDep_"+id).style.display="none";
		//document.getElementById("typeName"+"_"+id).style.display="block";

		}
function OpenInput(n,value)
{
$('form input').hide();
 $('form select').hide();
$('div[id^="dvT"').show();

document.getElementById("dvT"+n+"_"+value).style.display="none";
document.getElementById("dv"+n+"_"+value).style.display="block";

document.getElementById(n+"_"+value).style.display = "block";

}
/*function OpenEmail(value)
{
$('form input').hide();
$('div[id^="dvT_"').show();

document.getElementById("dvT_"+value).style.display="none";
document.getElementById("dv_"+value).style.display="block";

document.getElementById("email_"+value).style.display = "block";

}*/
function edit_Selectrow(value)
{
 $('form select').hide();
 $('form input').hide();
 document.getElementById("dvTDep_"+value).style.display="none";
 document.getElementById("dvDep_"+value).style.display="block";
 document.getElementById("sel_"+value).style.display = "block";
///--get select dep
$.ajax({
    url : "CreateSelDepartments.aspx",
    type: "POST",
    data : "",
			success: function(data, textStatus, jqXHR)
			{
	//alert(data)
	$('#sel_'+value).children().remove().end().append(data) ;
	//	$("#sel_"+value).append(data);
		// var ResData new Array();
		// var ResData=data.split("_");
		//data - response from server
			},
			error: function (jqXHR, textStatus, errorThrown)
			{
		 
			}
});

 
 }
 function edit_row(value)
 {
 $('form select').hide();
 $('form input').hide();
document.getElementById("dvTDep_"+value).style.display="none";
document.getElementById("dvDep_"+value).style.display="block";
//document.getElementById("sel_"+value).style.display = "block";

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

    var table = $('#example').DataTable( {
    
        scrollY:        "80vh",
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
            leftColumns: 1,
            rightColumns: 1
        }
     } );

// alert(document.getElementById("dl").className);
 //$("#dl").removeClass("sorting").addClass("main");
 $(".dataTables_scrollBody").scrollLeft(9999)
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
        width: 100%;/*500px;/*70%;*/
        margin: 0 auto;
    }
</style>

  </head>
  <body >

    <form id="Form1" method="post" runat="server">
    
<table id="example" class="stripe row-border cell-border order-column" cellspacing="0" width="100%">
        <thead>
            <tr>
               
                <th  nowrap class="no-sort">עדכון</th>
    
                <asp:Repeater ID=rptTitles Runat=server>
                <ItemTemplate>
                 <th nowrap valign=middle ><b><%#Container.dataItem("parent_title")%></b><br><%#Container.DataItem("bar_title")%></th>
                 </ItemTemplate>
                </asp:Repeater>
         
                <th nowrap valign=top><B>בדש בורד</B><BR>צפיה איש קשר</th>  
                <th nowrap valign=top><B>טפסים וסקרים</B><BR>דוח מגמות ביקושים<BR>ואחוזי סגירה</th>
                <th nowrap valign=top><B>טפסים וסקרים</B><BR>דוח מעקב רישום</th>
                <th nowrap valign=top><b>אופרציה</b><BR>צפיה בכל הטיולים</th>
                <th nowrap valign=top><b>משובים</b><BR> שליחת סמס ממסך איש קשר</th>
                <th nowrap valign=top><b>משובים</b><BR>שליחת סמס ממסך מרכז</th>
                <th nowrap valign=top><B>מיקודים</B><BR>שינוי מיקום<BR>שורות במסך הצפיה</th>
                <th nowrap valign=top><B>מיקודים</B><br>שליחת אימייל<br> ממסך ההגדרות</th>
                <th nowrap valign=top><B>מיקודים</B><br>שינוי סטטוס</th>
                <th nowrap valign=top><B>מיקודים</B><br>תצוגת טיול</th>
                <th nowrap><B>דש בורד</B><br>שינוי מחלקה <BR>בטופס רישום חתום</th>
                <th nowrap>שליחת<br>אימייל קבוצתי </th>
                <th nowrap>שינוי תוכן<br>הודעת sms</th>
                <th nowrap>רישום לקבוצות<BR>הסגורות</th>
                <th nowrap>ביטוח <BR>דיוויד שילד</th>
                <th nowrap>מורשה העברה<br>טופס לארכיון <BR>/לטיפול חוזר</th>
                <th nowrap>מורשה<br>מחיקה</th>
                <th nowrap>כניסה<BR>IP-מ</th>
                <th nowrap>פעיל</th>
                <th nowrap>דרגת חשיבות<BR>ידיעת המשוב</th>
                <th>שווי כספי <BR> של הזמנה</th>
                <th>יעד הזמנות <BR>מינימלי</th>
                <th nowrap>מחלקה</th>
                <th nowrap>מייל</th>
                <th nowrap>סוג עובד</th>
                <th nowrap><B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;שם עובד&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</B></th>
            </tr>
        </thead>
        
        <tbody>
 	<asp:Repeater id="rptWorkers" runat="server" >

          <ItemTemplate> 
            <tr style="background-color:#f0f0f0" >
               
     
                 <asp:Repeater ID=rptData Runat=server>
              <ItemTemplate>
               <td><%#IIf (trim(func.DbNullFix(Container.DataItem("Is_Permision"))) = "1" ,"V","")%></td>
              </ItemTemplate>
              </asp:Repeater> 
              
            
                <td><%#IIf (trim(func.DbNullFix(Container.DataItem("DashBoardView"))) = "1" ,"V","")%></td>
                <td><%#IIf (trim(func.DbNullFix(Container.DataItem("ReportCloseProcView"))) = "1" ,"V","")%></td>
                <td><%#IIf (trim(func.DbNullFix(Container.DataItem("ReportMaakavRishumView"))) = "1" ,"V","")%></td>
                <td><%#IIf (trim(func.DbNullFix(Container.DataItem("User_OperationScreenView"))) = "1" ,"V","")%></td>
                <td><%#IIf (trim(Container.DataItem("ScreenContactSendSms")) = "1" ,"V","")%></td>
                <td><%#IIf (trim(Container.DataItem("ScreenGeneralSendSms")) = "1" ,"V","")%></td>
                <td><%#IIf (trim(Container.DataItem("ScreenTourOrder")) = "1" ,"V","")%></td>
                <td><%#IIf (trim(Container.DataItem("ScreenSendMail")) = "1" ,"V","")%></td>
                <td><%#IIf (trim(Container.DataItem("ScreenTourStatus")) = "1" ,"V","")%></td>
                <td><%#IIf (trim(Container.DataItem("Screen_TourVisible")) = "1" ,"V","")%></td>
                <td><%#IIf (trim(Container.DataItem("Edit_DepartmentAppeal")) = "1" ,"V","")%></td>
                <td><%#IIf (trim(Container.DataItem("EmailGroupSend")) = "1" ,"V","")%></td>
                <td><%#IIf (trim(Container.DataItem("Sms_Write")) = "1" ,"V","")%></td>
                <td><%#IIf (trim(Container.DataItem("Add_GroupsTours")) = "1" ,"V","")%></td>
                <td><%#IIf (trim(Container.DataItem("Add_Insurance")) = "1" ,"V","")%></td>
                <td><%#IIf (trim(func.dbNullFix(Container.DataItem("Archive_appeal"))) = "1" ,"V","")%></td>
              
                <td><%#IIf (trim(Container.DataItem("chief")) = "1" ,"V","")%></td>
                <td><%#IIf (trim(Container.DataItem("IP_login")) = "1" ,"V","")%></td>
                <td><%#IIf (Container.DataItem("active") = "0","<img src=../../images/lamp_off.gif border=0 WIDTH=13 HEIGHT=18>","<img src=../../images/lamp_on.gif  border=0 WIDTH=13 HEIGHT=18>")%></td>
                <td><asp:Label runat="server" ID="lblImpId" Text=""></asp:Label></td>
             <td align=center style="background-color:ffffff"  id="TdOrderPrice_<%#Container.DataItem("User_Id")%>" onclick="javascript:OpenInput('OrderPrice','<%#Container.DataItem("User_Id")%>')">
                <div id="dvOrderPrice_<%#Container.DataItem("User_Id")%>" style="display:none;">
                <input type=text dir="ltr" name="OrderPrice_<%#Container.DataItem("User_Id")%>" id="OrderPrice_<%#Container.DataItem("User_Id")%>" value="<%#Container.DataItem("Order_Price")%>"  style="width:60;font-family:arial" maxlength=100 onblur="javascript:changeInput('OrderPrice','<%#Container.DataItem("User_Id")%>');" >
                </div>
                <div id="dvTOrderPrice_<%#Container.DataItem("User_Id")%>"  style="display:block"><%#Container.DataItem("Order_Price")%></div>
                </td>      <td  style="background-color:ffffff" align=center id="TdminOrder_<%#Container.DataItem("User_Id")%>" onclick="javascript:OpenInput('minOrder','<%#Container.DataItem("User_Id")%>')">
                <div id="dvminOrder_<%#Container.DataItem("User_Id")%>" style="display:none;">
                <input type=text dir="ltr" name="minOrder_<%#Container.DataItem("User_Id")%>" id="minOrder_<%#Container.DataItem("User_Id")%>" value="<%#Container.DataItem("Month_Min_Order")%>"  style="width:60;font-family:arial" maxlength=100 onblur="javascript:changeInput('minOrder','<%#Container.DataItem("User_Id")%>');" >
                </div>
                <div id="dvTminOrder_<%#Container.DataItem("User_Id")%>"  style="display:block"><%#Container.DataItem("Month_Min_Order")%></div>
                </td>
           
               <td dir=rtl style="background-color:ffffff"  id="TdDep_<%#Container.DataItem("USER_ID")%>" onclick="edit_Selectrow('<%#Container.DataItem("USER_ID")%>')" align=center>
<div id="dvTDep_<%#Container.DataItem("USER_ID")%>" style="display:block">
<div id="typeName_<%#Container.DataItem("USER_ID")%>"><%#Container.DataItem("departmentName")%></div>
</div>
<div  id="dvDep_<%#Container.DataItem("USER_ID")%>" style="display:none">
<select id="sel_<%#Container.DataItem("USER_ID")%>" onchange="javascript:ChangeSelect('<%#Container.DataItem("USER_ID")%>');">
</select></div>
</td>
                <td style="background-color:ffffff"  id="Tdemail_<%#Container.DataItem("User_Id")%>" onclick="javascript:OpenInput('email','<%#Container.DataItem("User_Id")%>')">
                <div id="dvemail_<%#Container.DataItem("User_Id")%>" style="display:none">
                <input type=text dir="ltr" name="email_<%#Container.DataItem("User_Id")%>" id="email_<%#Container.DataItem("User_Id")%>" value="<%#Container.DataItem("email")%>"  style="width:200;font-family:arial"" maxlength=100 onblur="javascript:changeInput('email','<%#Container.DataItem("User_Id")%>');" >
                </div>
                <div id="dvTemail_<%#Container.DataItem("User_Id")%>"  style="display:block"><%#Container.DataItem("email")%></div>
                </td>
      <td><%#Container.DataItem("job_name")%></td>
                <td nowrap><b><%#Container.DataItem("user_name")%></b></td>
            </tr>
           </ItemTemplate>
          <FooterTemplate> </tbody></FooterTemplate> 
          </asp:Repeater>
       
    </table>
</form>

</body>

</html>
