<%Response.Expires=-1%>
<%Response.CharSet = 1255%>
<%
	lang_id	 = trim(Request.Cookies("bizpegasus")("LANGID"))
		
	If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 2
	End If
	If lang_id = 2 Then
		title = "Calendar"
	Else
		title = "לוח שנה"
	End If		

%>
<HTML>
   <HEAD>   
   <meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
      <TITLE><%=title%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TITLE>
      <STYLE TYPE="text/css">
         .today {
		 	color:#000000; 
		    BACKGROUND-COLOR: #c0e0f8;
		    BORDER-BOTTOM: #4880ab 1px inset;
		    BORDER-LEFT: #4880ab 1px solid;
		    BORDER-RIGHT: #4880ab 1px inset;
		    BORDER-TOP: #4880ab 1px solid;
		    FONT-FAMILY: Arial;
		    FONT-SIZE: 11px;
			}
         .days {
		    BACKGROUND-COLOR: #f1f1f1;
		    BORDER-BOTTOM: #949494 1px inset;
		    BORDER-LEFT: #b4b4b4 1px solid;
		    BORDER-RIGHT: #949494 1px inset;
		    BORDER-TOP: #b4b4b4 1px solid;
		    FONT-FAMILY: Arial;
		    FONT-SIZE: 11px;
		    FONT-WEIGHT: 400;
		    
		 }
		 .currday {
		    color:#000000; 
		    BACKGROUND-COLOR: #d0d0d0;
		    BORDER-BOTTOM: #000000 1px inset;
		    BORDER-LEFT: #000066 1px solid;
		    BORDER-RIGHT: #000066 1px inset;
		    BORDER-TOP: #000000 1px solid;
		    FONT-FAMILY: Arial;
		    FONT-SIZE: 11px;
		 
		 }
		 .dateSat {
		    BACKGROUND-COLOR: #f1f1f1;
		    BORDER-BOTTOM: #949494 1px inset;
		    BORDER-LEFT: #b4b4b4 1px solid;
		    BORDER-RIGHT: #949494 1px inset;
		    BORDER-TOP: #b4b4b4 1px solid;
		    FONT-FAMILY: Arial;
		    FONT-SIZE: 11px;
		    COLOR:#FF0000; 
		 }
		 .dateFri {
		    BACKGROUND-COLOR: #e8e8e8;
		    BORDER-LEFT: #f0f0f0 1px solid;
		    BORDER-RIGHT: #f0f0f0 1px solid;
		    BORDER-TOP: #f0f0f0 1px solid;
		    BORDER-BOTTOM: #f0f0f0 1px solid;
		    FONT-FAMILY: Arial;
		    FONT-SIZE: 11px;
		    FONT-WEIGHT: 400;
		 }
		.head {
		    BACKGROUND-COLOR: #cccccc;
		    BORDER-BOTTOM: #949494 1px inset;
		    BORDER-LEFT: #b4b4b4 1px solid;
		    BORDER-RIGHT: #949494 1px inset;
		    BORDER-TOP: #b4b4b4 1px solid;
		    FONT-FAMILY: Arial;
		    FONT-SIZE: 11px;
		    FONT-WEIGHT: 400;
		    
		 } 
		.dateNone {
		    BACKGROUND-COLOR: #e8e8e8;
		 } 
		.clsTabSelected {
		 BACKGROUND-COLOR: #e8e8e8;
		} 
		 .td_head_black
		{
		    COLOR: #000000;
		    FONT-FAMILY: Arial;
		    FONT-SIZE: 11px;
		    FONT-WEIGHT: 400;
		    TEXT-DECORATION: none; 
		    /* Cursor: hand; */
		}
		.form_txt_field
		{
		    BACKGROUND-COLOR: #ffffff;
		    FONT-FAMILY: Arial;
		    FONT-SIZE: 11px;
		    FONT-WEIGHT: 400;
		}
      </STYLE>
      
      <SCRIPT LANGUAGE="JavaScript">
        var dateFromClient,today, dateFromCal;
        var strArguments=new String();
            strArguments=window.dialogArguments;
        //alert(strArguments)
        if (strArguments.length > 0 ){
           dateFromClient=strArguments
           var arrDate=new Array(3)
           arrDate=dateFromClient.split("/")
           dateFromCal=new Date(arrDate[2],arrDate[1]-1,arrDate[0])
        }else{
           dateFromCal='';
        }  
        today = new getToday(dateFromCal); 
        var currday = new getToday('');

        // Initialize arrays.
		// Hebrew ----------------------------------------- 
		var daysInMonth = new Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);
		<%If trim(lang_id) = "1" Then%>
		var months = new Array("ינואר","פברואר","מרץ","אפריל",
											"מאי","יוני","יולי","אוגוסט",
											"ספטמבר","אוקטובר","נובמבר","דצמבר");
        var days = new Array(" א ", " ב ", " ג ", " ד ", " ה ", " ו ", " ש ");
		//------ ------------------------------------------- 
		<%Else%>
		// English -----------------------------------------
		    var months = new Array("January","February","March","April",
											"May","June","July","August",
											"September","October","November","December");
            var days = new Array(" Su ", " Mo ", " Tu ", " We ", " Th ", " Fr ", " Sa ");
        //------ ----------------------------------------- */         
        <%End If%>
         function getDays(month, year) {
            // Test for leap year when February is selected.
            if (1 == month)
               return ((0 == year % 4) && (0 != (year % 100))) ||
                  (0 == year % 400) ? 29 : 28;
            else
               return daysInMonth[month];
         }


       function getToday(fromClient) {
             // Generate today's date from client. 
         if (fromClient!=''){
             this.now = new Date(fromClient);
             this.year = this.now.getFullYear(fromClient);
             this.month = this.now.getMonth(fromClient);
             this.day = this.now.getDate(fromClient);
         }  
             // Generate today's date. 
         else{
             this.now = new Date();
             this.year = this.now.getFullYear();
             this.month = this.now.getMonth();
             this.day = this.now.getDate();
           }    
        }
         
        function newCalendar() {
            today = new getToday(dateFromCal);
            var parseYear = parseInt(document.all.year[document.all.year.selectedIndex].text);
            var newCal = new Date(parseYear,document.all.month.selectedIndex, 1);
            var day = -1;
            var startDay = newCal.getDay();
            var daily = 0;
            if ((today.year == newCal.getFullYear()) && (today.month == newCal.getMonth()))
               day = today.day;

            // Cache the calendar table's tBody section, dayList.
            var tableCal = document.all.calendar.tBodies.dayList;
            var intDaysInMonth = getDays(newCal.getMonth(), newCal.getFullYear());
            
            for (var intWeek = 0; intWeek < tableCal.rows.length; intWeek++)
               for (var intDay = 0; intDay < tableCal.rows[intWeek].cells.length; intDay++) {
                   var cell = tableCal.rows[intWeek].cells[intDay];
                   // Start counting days.
                   if ((intDay == startDay) && (0 == daily))
                      daily = 1;
                   // Highlight the current day.
                   //if  (intDay == 5){ 
                   //   cell.className = "dateFri";
                   //}else if (intDay == 6){
                   if (day == daily){
                      cell.className = "today";
                   }else if  (intDay == 6){    
                      cell.className = "dateSat";
                   }else{
                     cell.className = "days";
                   }
                   //cell.className = (day == daily) ? "today" : "";

                   // Output the day number into the cell.
                   if ((daily > 0) && (daily <= intDaysInMonth))
                      cell.innerText = daily++;
                   else{
                      cell.className = "dateNone";
                      cell.innerText=""
                   }   
               }
         }

         function getDate() {
		 	var sDate;
            // This code executes when the user clicks on a day
            // in the calendar.
            if ("TD" == event.srcElement.tagName)
               // Test whether day is valid.
               if ("" != event.srcElement.innerText){
				  sDD  = event.srcElement.innerText
				  sDate =  sDD   + "/" + document.all.month.value + "/" + document.all.year.value;
	  		      strCurrDate= currday.day+'/'+ parseInt(currday.month+1) + '/' + currday.year
			      document.all.ret.value = sDate;
 		  		  window.close();
 		  	   }	  
         }
         
  </SCRIPT>
   </HEAD>
   <BODY ONLOAD="newCalendar()" OnUnload="window.returnValue = document.all.ret.value;" class="scrollbarDark" bgcolor="#e5e5e5" marginwidth="0" scroll="no" marginheight="0" hspace="0" vspace="0" topmargin="0" leftmargin="10" rightmargin="0">
     <input type="hidden" name="ret">
     <div align="center"> 
	 <TABLE ID="calendar" border="0" width="100%" align="center" class="td_head_black" cellspacing="1" cellpadding="1" dir="ltr">
            <TR ALIGN=CENTER>
               
               <TD COLSPAN=7 ALIGN=CENTER nowrap>
                  <!-- Month combo box -->
                  <SELECT ID="month" ONCHANGE="newCalendar()" Class=form_txt_field>
                     <SCRIPT LANGUAGE="JavaScript" dir="rtl">
                        // Output months into the document.
                        // Select current month.
                        for (var intLoop = 0; intLoop < months.length; intLoop++)
                           document.write("<OPTION  VALUE= " + (intLoop + 1) + " " + (today.month == intLoop ? "Selected" : "") + ">" + months[intLoop]);
                     </SCRIPT>
                  </SELECT>
                  <!-- Year combo box -->
                  <SELECT  ID="year" ONCHANGE="newCalendar()" Class=form_txt_field>
                     <SCRIPT LANGUAGE="JavaScript">
                        // Output years into the document.
                        // Select current year.
                        for (var intLoop = currday.year-100; intLoop < (currday.year + 10); intLoop++)
                           document.write("<OPTION VALUE= " + intLoop + " " + (today.year == intLoop ? "Selected" : "") + ">" + intLoop);
                     </SCRIPT>
                  </SELECT>
               </TD>
            </TR>
            
            <TR CLASS="head">
               <!-- Generate column for each day. -->
               <SCRIPT LANGUAGE="JavaScript">
                  // Output days.
                  for (var intLoop = 0; intLoop < days.length; intLoop++)
                     document.write("<TD align=center>" + days[intLoop] + "</TD>");
               </SCRIPT>
            </TR>
         <TBODY ID="dayList" ALIGN=CENTER>
            <!-- Generate grid for individual days. -->
            <SCRIPT LANGUAGE="JavaScript">
               for (var intWeeks = 0; intWeeks < 6; intWeeks++) {
                  document.write("<TR bgcolor=#ffffff>");
                  for (var intDays = 0; intDays < days.length; intDays++){
                     //if (intDays > 5)
                     //  document.write("<TD onclick='return false'></TD>");
                     //else 
                       document.write("<TD ONCLICK='getDate()'></TD>");
                  }    
                  document.write("</TR>");
               }
            </SCRIPT>
         </TBODY>
         <!--
         <tr><td colspan="7" height="2"></td></tr>
         <tr><td colspan="7" align="center"><Input class="buttons" style="font-size:11px;width:80px;"  type=button value="חזרה" OnClick="Cancel();" id=button1 name=button1></td></tr>
         <tr><td colspan="7" height="2"></td></tr>
         -->
      </TABLE> 
  </div>    
	  
<Script Language="JavaScript1.2">
	function Cancel() {
		document.all.ret.value = "";
		window.close();
	}
</script>
</body>
</html>