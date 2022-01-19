<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
Function CreateGUID()
  Randomize Timer
  Dim tmpCounter,tmpGUID
  Const strValid = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  For tmpCounter = 1 To 20
    tmpGUID = tmpGUID & Mid(strValid, Int(Rnd(1) * Len(strValid)) + 1, 1)
  Next
  CreateGUID = tmpGUID
End Function%>
<%

    If Request.QueryString("add") <> nil Then
     GUID=CreateGUID()
		If trim(Request.Form("supplier_Id")) = "" Then ' add type
			supplier_Id = trim(Request.Form("supplier_Id"))
		'	sqlstr = "Insert into Suppliers (supplier_Name,Country_Id) values ('" & sFix(Request.Form("supplier_Name")) & "',"& sFix(Request.Form("Country_Id")) &")"
sqlstr = "Insert into Suppliers (LOGINNAME,PASSWORD,supplier_Name,Country_Id,supplier_Name1,supplier_Name2,supplier_Name3,supplier_Name4,supplier_Job1,supplier_Job2,supplier_Job3,supplier_Job4," & _
" supplier_Tel1,supplier_Tel2,supplier_Tel3,supplier_Tel4, " & _
" supplier_Ext1,supplier_Ext2,supplier_Ext3,supplier_Ext4, " & _
" supplier_Phone1,supplier_Phone2,supplier_Phone3,supplier_Phone4, " & _
" supplier_Email1,supplier_Email2,supplier_Email3,supplier_Email4, " & _
" supplier_Descr1,supplier_Descr2,supplier_Descr3,supplier_Descr4, " & _
" TermsPayment,supplier_Descr,GUID) values ('" & sFix(Request.Form("LOGINNAME")) & "','"& sFix(Request.Form("PASSWORD")) &"','" & sFix(Request.Form("supplier_Name")) & "',"& sFix(Request.Form("Country_Id")) &",'"& sFix(Request.Form("supplier_Name1")) &"','"& sFix(Request.Form("supplier_Name2"))&"','"& sFix(Request.Form("supplier_Name3"))&"','"& sFix(Request.Form("supplier_Name4")) & _
"','"& sFix(Request.Form("supplier_Job1"))& "','" & sFix(Request.Form("supplier_Job2"))& "','" & sFix(Request.Form("supplier_Job3"))& "','" & sFix(Request.Form("supplier_Job4"))& "'," & _
"'"& sFix(Request.Form("supplier_Tel1"))& "','" & sFix(Request.Form("supplier_Tel2"))& "','" & sFix(Request.Form("supplier_Tel3"))& "','" & sFix(Request.Form("supplier_Tel4"))& "'," & _
"'"& sFix(Request.Form("supplier_Ext1"))& "','" & sFix(Request.Form("supplier_Ext2"))& "','" & sFix(Request.Form("supplier_Ext3"))& "','" & sFix(Request.Form("supplier_Ext4"))& "'," & _
"'"& sFix(Request.Form("supplier_Phone1"))& "','" & sFix(Request.Form("supplier_Phone2"))& "','" & sFix(Request.Form("supplier_Phone3"))& "','" & sFix(Request.Form("supplier_Phone4"))& "'," & _
"'"& sFix(Request.Form("supplier_Email1"))& "','" & sFix(Request.Form("supplier_Email2"))& "','" & sFix(Request.Form("supplier_Email3"))& "','" & sFix(Request.Form("supplier_Email4"))& "'," & _
"'"& sFix(Request.Form("supplier_Descr1"))& "','" & sFix(Request.Form("supplier_Descr2"))& "','" & sFix(Request.Form("supplier_Descr3"))& "','" & sFix(Request.Form("supplier_Descr4"))& "'," & _
"'"& sFix(Request.Form("TermsPayment"))&"','"& sFix(Request.Form("supplier_Descr")) & "','"& GUID &"')"
			
''response.Write sqlstr
''response.end
	
			con.executeQuery(sqlstr) %>
			<SCRIPT LANGUAGE=javascript>
			<!--
				window.opener.document.location.href = window.opener.document.location.href;
				window.close();
			//-->
			</SCRIPT>
		
	<%	Else ' update type
			supplier_Id = trim(Request.Form("supplier_Id"))
			sqlstr="Update Suppliers set supplier_Name = '" & sFix(Request.Form("supplier_Name")) & "',Country_Id=" &  sFix(Request.Form("Country_Id")) &" ,supplier_Name1 = '" & sFix(Request.Form("supplier_Name1")) &_
			"' ,supplier_Job1 = '" & sFix(Request.Form("supplier_Job1")) & _
			"' ,supplier_Tel1 = '" & sFix(Request.Form("supplier_Tel1")) & _
			"' ,supplier_Ext1 = '" & sFix(Request.Form("supplier_Ext1")) & _
			"' ,supplier_Phone1 = '" & sFix(Request.Form("supplier_Phone1")) & _
			"' ,supplier_Email1 = '" & sFix(Request.Form("supplier_Email1")) & _
			"' ,supplier_Descr1 = '" & sFix(Request.Form("supplier_Descr1")) & _
		    "' ,supplier_Name2 = '" & sFix(Request.Form("supplier_Name2")) & _
			"' ,supplier_Job2 = '" & sFix(Request.Form("supplier_Job2")) & _
			"' ,supplier_Tel2 = '" & sFix(Request.Form("supplier_Tel2")) & _
			"' ,supplier_Ext2 = '" & sFix(Request.Form("supplier_Ext2")) & _
			"' ,supplier_Phone2 = '" & sFix(Request.Form("supplier_Phone2")) & _
			"' ,supplier_Email2 = '" & sFix(Request.Form("supplier_Email2")) & _
			"' ,supplier_Descr2 = '" & sFix(Request.Form("supplier_Descr2")) & _
			"' ,supplier_Name3 = '" & sFix(Request.Form("supplier_Name3")) & _
			"' ,supplier_Job3 = '" & sFix(Request.Form("supplier_Job3")) & _
			"' ,supplier_Tel3 = '" & sFix(Request.Form("supplier_Tel3")) & _
			"' ,supplier_Ext3 = '" & sFix(Request.Form("supplier_Ext3")) & _
			"' ,supplier_Phone3 = '" & sFix(Request.Form("supplier_Phone3")) & _
			"' ,supplier_Email3 = '" & sFix(Request.Form("supplier_Email3")) & _
			"' ,supplier_Descr3 = '" & sFix(Request.Form("supplier_Descr3")) & _
	    	"' ,supplier_Name4 = '" & sFix(Request.Form("supplier_Name4")) & _
			"' ,supplier_Job4 = '" & sFix(Request.Form("supplier_Job4")) & _
			"' ,supplier_Tel4 = '" & sFix(Request.Form("supplier_Tel4")) & _
			"' ,supplier_Ext4 = '" & sFix(Request.Form("supplier_Ext4")) & _
			"' ,supplier_Phone4 = '" & sFix(Request.Form("supplier_Phone4")) & _
			"' ,supplier_Email4 = '" & sFix(Request.Form("supplier_Email4")) & _
			"' ,supplier_Descr4 = '" & sFix(Request.Form("supplier_Descr4")) & _
			"' ,TermsPayment = '" & sFix(Request.Form("TermsPayment")) & _
			"' ,supplier_Descr = '" & sFix(Request.Form("supplier_Descr")) & _
			"' ,LOGINNAME = '" & sFix(Request.Form("LOGINNAME")) & _
			"' ,PASSWORD = '" & sFix(Request.Form("PASSWORD")) & _
			"'  Where supplier_Id = " & supplier_Id
		
			con.executeQuery(sqlstr) %>
			
	<%	End If
	
	
	
		 %>
	<SCRIPT LANGUAGE=javascript>
			<!--
				window.opener.document.location.href = window.opener.document.location.href;
		        window.close();
			//-->
			</SCRIPT>	
	<%End If
	
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 33 Order By word_id"				
	  set rstitle = con.getRecordSet(sqlstr)		
	  If not rstitle.eof Then
	    dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(arr_title(0)) = arr_title(1)
			End If
		Next
	  End If
	  set rstitle = Nothing
	  
	  sqlstr = "select value From buttons Where lang_id = " & lang_id & " Order By button_id"
	  set rsbuttons = con.getRecordSet(sqlstr)
	  If not rsbuttons.eof Then
		arrButtons = ";" & rsbuttons.getString(,,",",";")		
		arrButtons = Split(arrButtons, ";")
	  End If
	  set rsbuttons=nothing	 	  
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
		function checkForm()
		{
			if(window.document.form1.supplier_Name.value == "")
			{
			   <%
				If trim(lang_id) = "1" Then
					str_alert = "!!נא להכניס ספק "
				Else
					str_alert = "Please insert the position name !!"
				End If   
				%>
				window.alert("<%=str_alert%>");				
				window.document.form1.supplier_Name.focus();
				return false;
			}
			return true;			
		}
//-->
</SCRIPT>
</HEAD>
<body style="margin:0px;background:#E5E5E5" onload="window.focus()">     
<%
	If Request.QueryString("supplier_Id") <> nil Then
		supplier_Id = trim(Request.QueryString("supplier_Id"))
		If Len(supplier_Id) > 0 Then
			sqlstr="Select supplier_Name,Country_Id,supplier_Name1,supplier_Name2,supplier_Name3,supplier_Name4,supplier_Job1," &_
			" supplier_Job2,supplier_Job3,supplier_Job4,supplier_Tel1,supplier_Tel2,supplier_Tel3,supplier_Tel4," &_
			" supplier_Ext1,supplier_Ext2,supplier_Ext3,supplier_Ext4," &_
			" supplier_Phone1,supplier_Phone2,supplier_Phone3,supplier_Phone4," &_
			" supplier_Email1,supplier_Email2,supplier_Email3,supplier_Email4," &_
			" supplier_Descr1,supplier_Descr2,supplier_Descr3,supplier_Descr4," &_
		    " TermsPayment,supplier_Descr,LOGINNAME,PASSWORD From Suppliers Where supplier_Id = " & supplier_Id
			set rssub = con.getRecordSet(sqlstr)
			If not rssub.eof Then
				supplier_Name = trim(rssub("supplier_Name"))				
				Country_Id= trim(rssub("Country_Id"))	
				supplier_Name1= trim(rssub("supplier_Name1"))		
				supplier_Name2= trim(rssub("supplier_Name2"))		
				supplier_Name3= trim(rssub("supplier_Name3"))		
				supplier_Name4= trim(rssub("supplier_Name4"))	
				supplier_Job1= trim(rssub("supplier_Job1"))		
				supplier_Job2= trim(rssub("supplier_Job2"))		
				supplier_Job3= trim(rssub("supplier_Job3"))		
				supplier_Job4= trim(rssub("supplier_Job4"))	
				
				supplier_Tel1= trim(rssub("supplier_Tel1"))	
				supplier_Tel2= trim(rssub("supplier_Tel2"))	
				supplier_Tel3= trim(rssub("supplier_Tel3"))	
				supplier_Tel4= trim(rssub("supplier_Tel4"))	
				
				supplier_Ext1= trim(rssub("supplier_Ext1"))	
				supplier_Ext2= trim(rssub("supplier_Ext2"))	
				supplier_Ext3= trim(rssub("supplier_Ext3"))	
				supplier_Ext4= trim(rssub("supplier_Ext4"))	
				
				supplier_Phone1= trim(rssub("supplier_Phone1"))	
				supplier_Phone2= trim(rssub("supplier_Phone2"))	
				supplier_Phone3= trim(rssub("supplier_Phone3"))	
				supplier_Phone4= trim(rssub("supplier_Phone4"))	
				
				supplier_Email1= trim(rssub("supplier_Email1"))	
				supplier_Email2= trim(rssub("supplier_Email2"))	
				supplier_Email3= trim(rssub("supplier_Email3"))	
				supplier_Email4= trim(rssub("supplier_Email4"))	
				
				
				
				supplier_Descr1= trim(rssub("supplier_Descr1"))	
				supplier_Descr2= trim(rssub("supplier_Descr2"))	
				supplier_Descr3= trim(rssub("supplier_Descr3"))	
				supplier_Descr4= trim(rssub("supplier_Descr4"))	
				
				TermsPayment= trim(rssub("TermsPayment"))
				supplier_Descr= trim(rssub("supplier_Descr"))
				LOGINNAME= trim(rssub("LOGINNAME"))
				PASSWORD_U= trim(rssub("PASSWORD"))
			End If
			set rssub=Nothing
		End If
	End If	
%>
<table dir="<%=dir_var%>" border="0" width="480" cellspacing="0" cellpadding="0" align="center" ID="Table1">
<tr>
   <td  align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="1" cellspacing="0" ID="Table2">
	   <tr>		 
		 <td class="page_title" align="<%=align_var%>" valign="middle" width="100%"><%If Len(supplier_Id) > 0 Then%><span id=word1 name=word1><!--עדכון--><%=arrTitles(1)%></span><%Else%><span id="word2" name=word2><!--הוספת--><%=arrTitles(2)%></span><%End If%>&nbsp;<span id=word4 name=word4>ספק</span>&nbsp;</td>
	  </tr>
	 </table>
  </td>
</tr>
<tr><td bgcolor="#e6e6e6" background="../../images/Line_3D.gif" height=2 width="100%"></td></tr>
<tr><td height=5 nowrap></td></tr>
<tr><td width=100%>
<table width="480" cellspacing="1" cellpadding="1" align=center border="0" ID="Table3">
<form name=form1 id=form1 action="addSupplier.asp?add=1" target="_self" method="post">
<input type=hidden name=supplier_Id id=supplier_Id value="<%=supplier_Id%>">
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<input type="text" class="texts" name="supplier_Name" id="supplier_Name" value="<%=vFix(supplier_Name)%>" dir="<%=dir_obj_var%>" size=100 maxLength=100>	
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id=word3 name=word3><!--שם סדרה-->שם ספק</span></b>&nbsp;</td>	
</tr>
<TR>
	<td align="<%=align_var%>" width=330 nowrap>
	<select dir="rtl" name="Country_Id" class="norm" style="width:200px" ID="Country_Id">
	<option value="0" >בחר מדינה</option>

<%'set rs_Countries = con.GetRecordSet("Select Country_Id, Country_Name from pegasus.dbo.Countries order by Country_Name")
set rs_Countries = con.GetRecordSet("Select Country_Id, Country_Name from pegasus.dbo.Countries order by Country_Name")
while not rs_Countries.eof	%>
    		<option value="<%=trim(rs_Countries(0))%>" <%If trim(Country_Id) = trim(rs_Countries(0)) Then%> selected <%End If%>><%=rs_Countries(1)%></option>

<%	rs_Countries.moveNext
	wend
		set rs_Countries = nothing %>
	</select>
<td width="150" nowrap align="right">&nbsp;<b><span id="Span1" name=word3>מדינה בה נותן שירותים</span></b>&nbsp;</td>	
</tr>

<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="ltr" name="LOGINNAME" value="<%=vFix(LOGINNAME)%>"  style="width:200" maxlength=20 style="font-family:arial" ID="Text1"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;שם משתמש&nbsp;</td>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=password dir="ltr" name="PASSWORD" value="<%=vFix(PASSWORD_U)%>"  style="width:200" maxlength=20 style="font-family:arial" ID="Text2" readonly></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;סיסמה&nbsp;</td>
</tr>
<TR><td colspan=2 align="center">&nbsp;<b><span id="Span8" name=word4><!--שם סוג-->1 פרטי איש קשר </span></b>&nbsp;</td>	
</td>
</TR>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<input type="text" class="texts"  dir="<%=dir_obj_var%>"  name="supplier_Name1" id="supplier_Name1"  dir=rtl size=100  value="<%=vFix(supplier_Name1)%>">
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span2" name=word4><!--שם סוג--> שם </span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
<input type="text" class="texts"  dir="<%=dir_obj_var%>"  name="supplier_Job1" id="supplier_Job1"  dir=rtl size=100  value="<%=vFix(supplier_Job1)%>">
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span9" name=word4><!--שם סוג--> תפקיד </span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<input type="text" class="texts"  name="supplier_Ext1" id="supplier_Ext1"  size=20  value="<%=vFix(supplier_Ext1)%>" style="width:60px"> / שלוחה

<input type="text" class="texts"   name="supplier_Tel1" id="supplier_Tel1"  size=20  value="<%=vFix(supplier_Tel1)%>" style="width:100px">
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span10" name=word4><!--שם סוג--> טלפון משרד </span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
<input type="text" class="texts"   name="supplier_Phone1" id="supplier_Phone1"  size=100  value="<%=vFix(supplier_Phone1)%>">
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span11" name=word4><!--שם סוג-->נייד </span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
<input type="text" class="texts"    name="supplier_Email1" id="supplier_Email1"  size=100  value="<%=vFix(supplier_Email1)%>">
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span12" name=word4><!--שם סוג-->מייל </span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<textarea  dir="<%=dir_obj_var%>" class="texts" name="supplier_Descr1" id="supplier_Descr1"  size=100 rows=3 cols=150><%=vFix(supplier_Descr1)%></textarea>	
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span13" name=word4><!--שם סוג-->הערות</span></b>&nbsp;</td>	
</tr>
<TR><td colspan=2 align="center">&nbsp;<b><span id="Span14" name=word4><!--שם סוג-->2 פרטי איש קשר </span></b>&nbsp;</td>	
</td>
</TR>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<input type="text" class="texts"  dir="<%=dir_obj_var%>"  name="supplier_Name2" id="supplier_Name2"  dir=rtl size=100  value="<%=vFix(supplier_Name2)%>">
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span15" name=word4><!--שם סוג--> שם </span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
<input type="text" class="texts"  dir="<%=dir_obj_var%>"  name="supplier_Job2" id="supplier_Job2"  dir=rtl size=100  value="<%=vFix(supplier_Job2)%>">
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span16" name=word4><!--שם סוג--> תפקיד </span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<input type="text" class="texts"   name="supplier_Ext2" id="supplier_Ext2" size=20  value="<%=vFix(supplier_Ext2)%>" style="width:60px"> / שלוחה

<input type="text" class="texts"    name="supplier_Tel2" id="supplier_Tel2"  size=20  value="<%=vFix(supplier_Tel2)%>" style="width:100px">
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span17" name=word4><!--שם סוג--> טלפון משרד </span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
<input type="text" class="texts"   name="supplier_Phone2" id="supplier_Phone2"   size=100  value="<%=vFix(supplier_Phone2)%>">
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span18" name=word4><!--שם סוג-->נייד </span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
<input type="text" class="texts"    name="supplier_Email2" id="supplier_Email2"   size=100  value="<%=vFix(supplier_Email2)%>">
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span19" name=word4><!--שם סוג-->מייל </span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<textarea  dir="<%=dir_obj_var%>" class="texts" name="supplier_Descr2" id="supplier_Descr2"  dir=rtl size=100 rows=3 cols=150><%=vFix(supplier_Descr2)%></textarea>	
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span20" name=word4><!--שם סוג-->הערות</span></b>&nbsp;</td>	
</tr>

<TR><td colspan=2 align="center">&nbsp;<b><span id="Span21" name=word4><!--שם סוג-->3 פרטי איש קשר </span></b>&nbsp;</td>	
</td>
</TR>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<input type="text" class="texts"  dir="<%=dir_obj_var%>"  name="supplier_Name3" id="supplier_Name3"  dir=rtl size=100  value="<%=vFix(supplier_Name3)%>">
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span22" name=word4><!--שם סוג--> שם </span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
<input type="text" class="texts"  dir="<%=dir_obj_var%>"  name="supplier_Job3" id="supplier_Job3"  dir=rtl size=100  value="<%=vFix(supplier_Job3)%>">
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span23" name=word4><!--שם סוג--> תפקיד </span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<input type="text" class="texts"   name="supplier_Ext3" id="supplier_Ext3"  size=20  value="<%=vFix(supplier_Ext3)%>" style="width:60px"> / שלוחה

<input type="text" class="texts"   name="supplier_Tel3" id="supplier_Tel3"  size=20  value="<%=vFix(supplier_Tel3)%>" style="width:100px">
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span24" name=word4><!--שם סוג--> טלפון משרד </span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
<input type="text" class="texts"    name="supplier_Phone3" id="supplier_Phone3"  size=100  value="<%=vFix(supplier_Phone3)%>">
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span25" name=word4><!--שם סוג-->נייד </span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
<input type="text" class="texts"  name="supplier_Email3" id="supplier_Email3"   size=100  value="<%=vFix(supplier_Email3)%>">
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span26" name=word4><!--שם סוג-->מייל </span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<textarea  dir="<%=dir_obj_var%>" class="texts" name="supplier_Descr3" id="supplier_Descr3"  dir=rtl size=100 rows=3 cols=150><%=vFix(supplier_Descr3)%></textarea>	
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span27" name=word4><!--שם סוג-->הערות</span></b>&nbsp;</td>	
</tr>
<TR><td colspan=2 align="center">&nbsp;<b><span id="Span28" name=word4><!--שם סוג-->4 פרטי איש קשר </span></b>&nbsp;</td>	
</td>
</TR>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<input type="text" class="texts"  dir="<%=dir_obj_var%>"  name="supplier_Name4" id="supplier_Name4"  dir=rtl size=100  value="<%=vFix(supplier_Name4)%>">
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span29" name=word4><!--שם סוג--> שם </span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
<input type="text" class="texts"  dir="<%=dir_obj_var%>"  name="supplier_Job4" id="supplier_Job4"  dir=rtl size=100  value="<%=vFix(supplier_Job4)%>">
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span30" name=word4><!--שם סוג--> תפקיד </span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<input type="text" class="texts"   name="supplier_Ext4" id="supplier_Ext4"  size=20  value="<%=vFix(supplier_Ext4)%>" style="width:60px"> / שלוחה

<input type="text" class="texts"    name="supplier_Tel4" id="supplier_Tel4"  size=20  value="<%=vFix(supplier_Tel4)%>" style="width:100px">
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span31" name=word4><!--שם סוג--> טלפון משרד </span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
<input type="text" class="texts"    name="supplier_Phone4" id="supplier_Phone4"  size=100  value="<%=vFix(supplier_Phone4)%>">
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span32" name=word4><!--שם סוג-->נייד </span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
<input type="text" class="texts"  name="supplier_Email4" id="supplier_Email4"   size=100  value="<%=vFix(supplier_Email4)%>">
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span33" name=word4><!--שם סוג-->מייל </span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<textarea   class="texts" name="supplier_Descr4" id="supplier_Descr4"  dir=rtl size=100 rows=3 cols=150><%=vFix(supplier_Descr4)%></textarea>	
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id="Span34" name=word4><!--שם סוג-->הערות</span></b>&nbsp;</td>	
</tr>
 <tr><td colspan=2 height=15></td></tr>

<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<textarea  dir="<%=dir_obj_var%>" class="texts" name="TermsPayment" id="TermsPayment"  dir=rtl size=100 rows=5 cols=150><%=vFix(TermsPayment)%></textarea>	
	</td>
	<td width="150" nowrap align="right" valign=top>&nbsp;<b><span id="Span6" name=word4><!--שם סוג-->תנאי תשלום</span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<textarea  dir="<%=dir_obj_var%>" class="texts" name="supplier_Descr" id="supplier_Descr"  dir=rtl size=100 rows=5 cols=150><%=vFix(supplier_Descr)%></textarea>	
	</td>
	<td width="150" nowrap align="right" valign=top>&nbsp;<b><span id="Span7" name=word4><!--שם סוג-->הערות</span></b>&nbsp;</td>	
</tr>



</table>

</td></TR>

<tr><td height=5 colspan="2" nowrap></td></tr>
<tr><td align=center colspan="2">
<input type=button value="<%=arrButtons(2)%>" class="but_menu" style="width:90" onclick="window.close();" id=Button2 name=Button2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type=submit value="<%=arrButtons(1)%>" class="but_menu" style="width:90" onclick="return checkForm()" id=Button1 name=Button1>
</td></tr>
<tr><td height=65 colspan="2" nowrap></td></tr>

</table>

</form>
</td></tr></table>
</BODY>
</HTML>
