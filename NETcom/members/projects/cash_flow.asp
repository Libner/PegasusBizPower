<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 62 Order By word_id"				
	  set rstitle = con.getRecordSet(sqlstr)		
	  If not rstitle.eof Then
	    dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(trim(arr_title(0))) = arr_title(1)
			End If
		Next
	  End If
	  set rstitle = Nothing		  

%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-type_id content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type_id="text/css">
<script LANGUAGE="JavaScript">
<!--	
	function report_open(fname)
	{	
		formsearch.action=fname;
		formsearch.target = '_blank';
		formsearch.submit();	
    }
    
	function popupcal(elTarget)
	{
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("../../calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=134pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		return false;
		window.document.focus;   
   }

   function GetNumbers()
   {
		var ch=event.keyCode;
		event.returnValue =(ch >= 48 && ch <= 57) || ch == 46 || ch == 45;
   } 

//-->	
</script> 
</head> 
<%
start_date = DateValue("1/" & Month(Date()) & "/" & Year(date()))
end_date = DateAdd("d",30,start_date)
balance = 0
%>
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 3%>
<%numOfLink = 6%>
<!--#include file="../../top_in.asp"-->
<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">
<tr>
   <td bgcolor="#e6e6e6" align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0" ID="Table8">
	  <tr><td class="page_title">&nbsp;</td></tr>		   
	  		       	
   </table></td></tr>         
<tr><td width=100%>
<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">
<tr><td align="<%=align_var%>" height="30" nowrap></td></tr>  
   <tr>    
    <td width="100%" valign="top" align="center">
	<FORM method=POST id=formsearch name=formsearch>   
    <table cellspacing=1 width=400 align=center border=0 bgcolor="#646E77" cellpadding="4">  
    <tr>		
		<td align="<%=align_var%>" bgcolor="#DBDBDB" width=100%>
		<input dir="ltr" class="texts" type="text" id="start_date" name="start_date" value="<%=start_date%>" style="width:75" onclick="return popupcal(this);" readonly>
		&nbsp;&nbsp;
		</td>		
		<td align="<%=align_var%>" width=100 class="subject_form" nowrap><b><span id=word1 name=word1><!--מתאריך--><%=arrTitles(1)%></span></b>&nbsp;</td>					
	</tr>	
    <tr>    	
		<td align="<%=align_var%>" bgcolor="#DBDBDB">
		<input dir="ltr" class="texts" type="text" id="end_date" name="end_date" value="<%=end_date%>" style="width:75" onclick="return popupcal(this);" readonly>
		&nbsp;&nbsp;
		</td>		
		<td align="<%=align_var%>" class="subject_form">&nbsp;<b><span id="word2" name=word2><!--עד תאריך--><%=arrTitles(2)%></span></b>&nbsp;</td>
	</tr>		      
   <tr><td align="<%=align_var%>" bgcolor="#DBDBDB">
   <select dir="<%=dir_obj_var%>" name="bank_id" style="width:100px;" class="norm" ID="bank_id">		
	<%  sqlstr = "Select bank_id, bank_name FROM banks WHERE ORGANIZATION_ID = " & OrgID & " Order BY bank_name"
		set rs_bank = con.getRecordSet(sqlstr)
		While not rs_bank.eof
	%>
	<option value="<%=rs_bank(0)%>" <%If trim(bank_id) = trim(rs_bank(0)) Then%> selected <%End If%>><%=rs_bank(1)%></option>
	<%
		rs_bank.moveNext
		Wend
		set rs_bank = Nothing
	%>
	</select>
	&nbsp;&nbsp;
	</td>
	<td width=100% align="<%=align_var%>" class="subject_form">&nbsp;<b><span id="word3" name=word3><!--חשבון בנק--><%=arrTitles(3)%></span></b>&nbsp;</td>
	</tr>
	<tr>    	
		<td align="<%=align_var%>" bgcolor="#DBDBDB">
		<input dir="ltr" class="texts" type="text" id="balance" name="balance" value="<%=balance%>" style="width:70" onkeypress="GetNumbers()">
		&nbsp;&nbsp;
		</td>		
		<td width=100% align="<%=align_var%>" class="subject_form">&nbsp;<b><span id="word4" name=word4><!--יתרה--><%=arrTitles(4)%></span></b>&nbsp;</td>
	</tr>
    <tr><td align="<%=align_var%>" bgcolor="#DBDBDB">
    <select dir="<%=dir_obj_var%>" name="type_id" style="width:90px;" class="norm" ID="type_id">		
	<option value="1" <%If trim(type_id) = trim("1") Then%> selected <%End If%> id="word5" name=word5><!--תכנון--><%=arrTitles(5)%></option>
	<option value="2" <%If trim(type_id) = trim("2") Then%> selected <%End If%> id="word6" name=word6><!--ביצוע--><%=arrTitles(6)%></option>
	</select>
	&nbsp;&nbsp;
	</td>
	<td width=100% align="<%=align_var%>" class="subject_form">&nbsp;<b><span id="word7" name=word7><!--סוג תזרים--><%=arrTitles(7)%></span></b>&nbsp;</td>
	</tr>		  
   <tr><td nowrap colspan=2 bgcolor="#DBDBDB" align="center"><a class="button_edit_1" style="width:100;" href="#" onclick="report_open('report_cash_flow.asp');">&nbsp;<span id="word8" name=word8><!--הצג תזרים--><%=arrTitles(8)%></span>&nbsp;</a></td></tr>
   </form>   
</table></td></tr></table>
</td></tr></table>
</td></tr></table>
</body>
</html>
<%
set con = nothing
%>