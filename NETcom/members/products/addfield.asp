<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="JavaScript">
function CheckDel() {
  <%
     If trim(lang_id) = "1" Then
        str_confirm = "?האם ברצונך למחוק"
     Else
		str_confirm = "Are you sure want to delete?"
     End If   
  %>  
   return (window.confirm("<%=str_confirm%>"))    
 }

function CheckFields()
{
	  var doc=window.document;	 
	  if (isNaN(document.forms[0].size.value))
	     {alert('שדה מספר');
	      document.forms[0].size.focus();
	      return false;
	     }
	  if (document.forms[0].type_field.value=='6')
	     { document.forms[0].size.value='10'}
}
function  chanFieldsObj(field_type,imgObj)
{
	 if (field_type=='7')
	 { document.forms[0].size.value='10'}	 
	 else 
	   {document.forms[0].size.value='50' }
	   
	 if (field_type=='4' || field_type=='9' || field_type=='12'){
	 	scale_tr.style.display = 'block';
	 }else{
	 	scale_tr.style.display = 'none';
	 }
	 
	 if (field_type=='1' || field_type=='7')
	 { size_tr.style.display = 'block'}
	 else 
	   {size_tr.style.display = 'none'}
	   
	 if (field_type=='1' || field_type=='2' || field_type=='3')
	 { align_tr.style.display = 'block'}
	 else 
	   {align_tr.style.display = 'none'}
	 
	 if ((field_type=='3' || field_type=='8' || field_type=='11' ) && (typeof(options_tr) != 'undefined'))
	 { options_tr.style.display = 'block'}
	 else if(typeof(options_tr) != 'undefined')
	 {options_tr.style.display = 'none'}  
	   
	window.document.all("type_field").value = field_type;
	   
	fieldCollection = document.getElementsByName("FIELD");
	//window.alert(fieldCollection.length);
	for(i=0;i<fieldCollection.length;i++)
	{
		if(fieldCollection(i).style.border && (fieldCollection(i).name == "FIELD"))
			fieldCollection(i).style.border = "2px solid #C0C0C0";
	}
	bordSt = new String(imgObj.style.border);
	//window.alert(bordSt.search("2px"))
	if(bordSt.search("2px") > 0)
		imgObj.style.border = "3px solid Red"
	else	
		imgObj.style.border = "2px solid #C0C0C0";	 
}

function changeDir(dir,imgObj)
{
	window.document.all("type_align").value = dir;
	fieldCollection = document.getElementsByName("DIR");
	//window.alert(fieldCollection.length);
	for(i=0;i<fieldCollection.length;i++)
	{
		if(fieldCollection(i).style.border)
			fieldCollection(i).style.border = "2px solid #C0C0C0";
	}
	
	bordSt = new String(imgObj.style.border);
	//window.alert(bordSt.search("2px"))
	if(bordSt.search("2px") > 0)
		imgObj.style.border = "3px solid Red"
	else	
		imgObj.style.border = "2px solid #C0C0C0";
	return false;		
}
	 
</script>

<style type="text/css">
    #divText {position:absolute;  left:120; top:400; display: none; z-index:4;}
</style>

<script>
function onSubmit_Text()
{
	if (window.document.form1.textItem.value=='')
	{
	  <%
		If trim(lang_id) = "1" Then
			str_alert = "נא להכניס אפשרות בחירה"
		Else
			str_alert = "Please insert the options"
		End If   
	%>  
		window.alert("<%=str_alert%>");
	}
	else
	{
		window.document.form1.submit()
	}
	return false		
}


function initcontent()
{
   document.all["divText"].style.display="none";
}

function visFormElem(Text,id)
{
    var offset = 250;
    scr_height = window.screen.availHeight;
	offset = parseInt(scr_height / 2);
	var answ_Rev=new String(Text);  
    document.all["divText"].style.zIndex=10 ; 
	document.all["divText"].style.display="inline"; 
	document.all["divText"].style.pixelTop=offset+document.body.scrollTop;	
		
	if (Text!='')
	{
		for (j=0;j<Text.length;j++)
		{
			 temp=answ_Rev.replace("<br>","\n")
			 answ_Rev=temp;
		}
		window.document.form1.idInf.value=id;
		window.document.form1.textItem.value=answ_Rev;	
		 <%
			If trim(lang_id) = "1" Then
				str_html = "<b>עדכן אפשרות בחירה</b>"
			Else
				str_html = "<b>Edit the option</b>"
			End If   
		%>  	
		window.titleEd.innerHTML="<%=str_html%>";		
		window.divText.style.pixelTop=window.divText.style.pixelTop-40;		
	} 
	else
	{
		window.document.form1.idInf.value=id;
		window.document.form1.textItem.value='';
		 <%
			If trim(lang_id) = "1" Then
				str_html = "<b>הוספת אפשרות</b>"
			Else
				str_html = "<b>Add an option</b>"
			End If   
		%>  			
		window.titleEd.innerHTML='<%=str_html%>'; 
	}
	window.document.form1.textItem.select();
//	return false;
}
 
function hideFormElem(argType) 
{ 
  var elementHid=argType; 
  document.all["divText"].style.display="none"; 
  return false;
} 
 
</script>
</head>

<%
prodId=trim(Request("prodId"))
id=trim(Request("Id"))

If trim(prodId) <> "" And trim(id) <> ""  Then
    sqlstr = "Select TOP 1 QUESTIONS_ID From APPEALS WHERE QUESTIONS_ID = " & prodId
    set rs_check = con.getRecordSet(sqlstr)
    If not rs_check.eof Then
    	is_send = true
    Else
    	is_send = false
    End If	
    set rs_check = Nothing
Else
	is_send = false    
End If

found_product = false		
If IsNumeric(prodId) = true And IsNull(prodId) = false Then
'Response.Write "SELECT Product_Name,Langu FROM Products where Product_Id="&prodId&" and ORGANIZATION_ID=" & trim(OrgID)
'Response.End
	Set product=con.GetRecordSet("SELECT Product_Name,Langu FROM Products where Product_Id="&prodId&" and ORGANIZATION_ID=" & trim(OrgID) )
    if not product.Eof then	
		productName=product("Product_Name")
		if product("Langu") = "eng" then
			dir_align = "ltr"
			td_align = "left"
			pr_language = "eng"
			p_align = "ltr"
		else
			dir_align = "rtl"
			td_align = "right"
			pr_language = "heb"
			p_align = "rtl"
		end if
		found_product = true
	else
	    found_product = false
	end if 'not product.Eof
	Set product= nothing
End If

sub set_fields_order(prodId)
'order
strTmp="SELECT Field_Order,Field_ID FROM Form_Field WHERE Product_ID = " & prodId & " and ORGANIZATION_ID=" & trim(OrgID) & " ORDER BY Field_Order "
set rsTmp=con.GetRecordSet(strTmp)
			cnt=rsTmp.RecordCount 
			if not rsTmp.EOF then
			 pp=1
				do while pp<=cnt
				 pgId=rsTmp("Field_ID")
				 stst="update Form_Field set Field_Order=" & pp & " where Field_ID=" & pgId
				 con.ExecuteQuery(stst)
			 pp=pp+1
			 rsTmp.MoveNext 
				loop
			end if
	set rsTmp=nothing
'end order
end sub

'delete  SELECT Value
if request("idSelect")<>nil and request("selectdel")<>nil and IsNumeric(prodId) = true And IsNull(prodId) = false Then
str_Delete="Delete from Form_Select where  Id=" & request("IdSelect") & " and ORGANIZATION_ID=" & trim(OrgID)
con.ExecuteQuery(str_Delete)
%>
 <script>
     window.document.location.href =  "addfield.asp?ID=<%=ID%>&prodID=<%=prodID%>#options_table";
 </script>
<%
end if

'end delete  SELECT Value


'add update SELECT Value
if  Request.Form("textItem")<>nil then
		textItem=sFix(trim(Request.Form("textItem")))
		idinf=Request.Form("idinf")
		if idInf<>0 then
			  textItem=sFix(trim(Request.Form("textItem")))
			  strUpd="UPDATE Form_Select SET Field_Value='"& textItem &"' WHERE Id=" & idInf & " and ORGANIZATION_ID=" & trim(OrgID)
			  con.ExecuteQuery (strUpd)
		 else
  		      strInsert="INSERT INTO Form_Select (Field_Id,ORGANIZATION_ID,Field_Value) VALUES ("& id &","& trim(OrgID) &",'" & textItem & "')"
		      'Response.Write(strInsert)		      
		      con.ExecuteQuery (strInsert)
         end if
         %>
         <script>
         window.document.location.href = "addfield.asp?ID=<%=ID%>&prodID=<%=prodID%>#options_table";
         </script>
         <%
        
 end if  
'end 

if request.form("type_field") = "4" or request.form("type_field") = "9" or request.form("type_field") = "12" then ' radio
	f_scale = request.form("selectscale")
	f_exception = ""
else
	f_scale = "null"
	f_exception = ""
end if

if request.form("type_field") = "1" or request.form("type_field") = "6" or request.form("type_field") = "7" then
	f_size = request.form("size")
else
	f_size = "null"
end if

If trim(Request.Form("must")) = nil Then
	fieldMust = "0"
Else
	fieldMust = "1"
End If		

if Request.Form("update")<>nil and id<>nil  then

   str_Update="UPDATE Form_Field SET Field_Title='"& sFix(trim(request.form("title"))) & "' , Field_Align = '"& request.form("type_align") & "',Field_Size= " & f_size& ",Field_Type=  " & request.form("type_field") & ",Field_Scale=  " & f_scale & " ,Field_Must = " & fieldMust & "  WHERE Field_id=" & id & " and ORGANIZATION_ID=" & trim(OrgID)
   'Response.Write(str_Update)
   'Response.End 
   con.ExecuteQuery(str_Update)
   'Response.Write prodID
   'Response.End
   set_fields_order(prodID)
  %>	
	<SCRIPT LANGUAGE=javascript>
	<!--
		document.location.href='addform.asp?prodId=<%=prodid%>#link<%=prodID%>';
	//-->
	</SCRIPT>
  <%		
  end if
if request.form("new_new")<>nil and id=nil and prodId <> nil then
       sqlstring="insert into Form_Field (Product_ID,ORGANIZATION_ID,Field_Title,Field_Type,Field_Size,Field_Align,Field_Order,Field_Scale,Field_Must) values (" &prodId & ","& trim(OrgID) &",'" & sFix(trim(request.form("title"))) & "'," & request.form("type_field") & "," & f_size& ",'" & sFix(request.form("type_align")) & "',1000," & f_scale &"," & fieldMust & ")"
       'Response.Write(sqlstring)
       'Response.end
       con.ExecuteQuery(sqlstring)
       set_fields_order(prodID)
      set pr = con.GetRecordSet("select Field_ID from Form_Field  WHERE ORGANIZATION_ID=" & trim(OrgID) & " ORDER BY Field_ID DESC")
		Id=trim(pr("Field_ID"))
      set pr = nothing
      if request.form("type_field")<>"3" and request.form("type_field")<>"8" and request.form("type_field")<>"11" then %>
      <SCRIPT LANGUAGE=javascript>
		<!--
		document.location.href='addform.asp?prodId=<%=prodid%>#link<%=prodID%>';
		//-->
	  </SCRIPT>
	  <%
	  else
	  %>
	  <SCRIPT LANGUAGE=javascript>
		<!--
		document.location.href='addfield.asp?ID=<%=ID%>&prodID=<%=prodID%>&viewPopup=1';
		//-->
	  </SCRIPT>
	   
<%	  end if
end if%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 22 Order By word_id"				
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
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 1%>
<%numOfLink = 1%>
<%topLevel2 = 14 'current bar ID in top submenu - added 03/10/2019%>
<%If trim(wizard_id) = "" then%>
<!--#include file="../../top_in.asp"-->
<%End if%>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">
<tr>
<td class="page_title" dir="<%=dir_obj_var%>"><%If IsNumeric(wizard_id) Then%> <%=page_title_%> <%Else%> <%=ProductName%> <%End If%>&nbsp;</td>
</tr>
<%If trim(wizard_id) <> "" Then%>
<tr><td width=100% align="<%=align_var%>">
<% wizard_page_id = 1 %>
<!--#include file="../../wizard_inc.asp"-->
</td></tr>
<tr>
<td width=100% align="<%=align_var%>" bgcolor="#FFD011" style="padding:5px">
<table border=0 cellpadding=0 cellspacing=0 width=100% ID="Table4">
<tr><td class="explain_title">הוסף את שדות הטופס</td></tr>
<tr><td height=5 nowrap></td></tr>
<tr><td class=explain>
בשורה <b>כותרת שדה</b> יש  להקליד את השאלה, ומתחתיה לבחור באיזו תבנית תהייה השאלה.
</td></tr>
<tr><td class=explain>
לאחר שבחרת תבנית, יופיעו שדות רלוונטיים לתבנית השאלה שבחרת.
</td></tr>
<tr><td class=explain>
עליך למלא את הפרטים הרצויים לך, למשל כיוון המלל בשדה השאלה וגודל שדה מקסימלי.
</td></tr>
<tr><td class=explain>
לדוגמה, בתבנית  "אפשרויות בחירה", תבקש ממך המערכת להוסיף את אפשרויות הבחירה של השדה.
</td></tr>
</table>
<tr>
    <td width="100%" height="2"></td>
</tr>
<%End if%>
<%If found_product Then%>
<tr>
<td align="<%=align_var%>" colspan=3>
<table cellpadding=0 cellspacing=0 width=100% border=0 ID="Table2">
<tr>
<td width=100%>
<table border="0" cellspacing="1" cellpadding="1" align="center" width="600" bgcolor="#ffffff">

<% p_scale=null
   p_size=30 
   p_type = "1"
   
If id <>  nil Then
'Response.Write "select * from Form_Field where Field_Id ="& id &" and ORGANIZATION_ID=" & trim(OrgID)
'Response.End
   set p_field=con.GetRecordSet("select * from Form_Field where Field_Id ="& id &" and ORGANIZATION_ID=" & trim(OrgID) )
   if not p_field.eof then
      p_title=p_field("Field_Title")  
      p_type=p_field("Field_Type")
      p_size=p_field("Field_Size")
      p_align=p_field("Field_Align")
      p_scale=p_field("Field_Scale")
      p_exception=p_field("FIELD_EXCEPTION")
      p_must = trim(p_field("Field_MUST"))	  
   end if 
 end if
	if p_type = "1" or p_type = "7" or id=nil then
		showSize = true
	else
		showSize = false
	end if
	if p_type = "1" or p_type = "2" or p_type = "3"  or id=nil then
		showAlign = true
	else
		showAlign = false
	end if
	if not IsNull(p_exception) or p_exception <> "" then
		myarr = Split(p_exception)
		if ubound(myarr) > 0 then
			p_condition = myarr(0)
			p_exception = myarr(1)
		end if	
	end if
		
%>
<form method="post" action="addfield.asp" id=formField name=formField> 
<tr><td width=100% colspan=2>
<table cellpadding=0 cellspacing=0 width=100% dir="<%=dir_var%>" border=0>
<tr>
	<td align="<%=align_var%>"><input class="passw"  dir="<%=dir_obj_var%>" name="Title" style="width:350" value="<%=vfix(p_title)%>"></td>
	<th class="title_show_form" width=150  align="<%=align_var%>" nowrap><span id=word3 name=word3><!--כותר שדה--><%=arrTitles(3)%></span>&nbsp;</th>
</tr>	
</table></td>	
</tr>
<tr>
	<td align="<%=align_var%>" colspan=2>	 
	 <!--select name="type_field" class="norm"  <%if id<>nil and (p_type="3" or p_type="8") then%>disabled<%else%>onchange="chanFieldsObj()"<%end if%>>
	   <option value="1" <%if p_type="1" then%> selected<%end if%>>שדה פתוח (שורה אחת)</OPTION>
	   <option value="2" <%if p_type="2" then%> selected<%end if%>>שדה פתוח (טקסט)</OPTION>
	   <option value="3" <%if p_type="3" then%> selected<%end if%>> תיבת רשימה - Combo</OPTION> 
	   <option value="4" <%if p_type="4" then%> selected<%end if%>>שאלת דירוג משוקללת</OPTION> 
	   <option value="9" <%if p_type="9" then%> selected<%end if%>>שאלת דירוג</OPTION>
	   <option value="8" <%if p_type="8" then%> selected<%end if%>> בחירה - Radio Button</OPTION>
	   <option value="5" <%if p_type="5" then%> selected<%end if%>>תיבת סימון</OPTION> 
	   <option value="6" <%if p_type="6" then%> selected<%end if%>>תאריך</OPTION>
	   <option value="7" <%if p_type="7" then%> selected<%end if%>>מספר</OPTION>
	 </select-->  
	 <%
		If trim(lang_id) = "1" Then
			img_ = ""
		Else
			img_ = "_eng"
		End If
	 %> 	
	 <table cellpadding=1 cellspacing=8 align="<%=align_var%>">
	 <tr>
	 <td name="FIELD" valign="top">
	 <img src="../../images/combo<%=img_%>.gif" <%If trim(p_type) = "3" Then%> style="border: 3px solid red" <%Else%> style="border: 2px solid #C0C0C0" <%End If%> <%If is_send = false Then%> class=hand onClick=" chanFieldsObj(3,this)" <%End If%> name="FIELD">
	 </td>
	 <td name="FIELD" valign="top">
	 <img src="../../images/textarea<%=img_%>.gif" <%If trim(p_type) = "2" Then%> style="border: 3px solid red" <%Else%> style="border: 2px solid #C0C0C0" <%End If%> <%If is_send = false Then%> class=hand onClick=" chanFieldsObj(2,this)" <%End If%> name="FIELD"></td>
	 <td name="FIELD" valign="top">
	 <img src="../../images/input<%=img_%>.gif" <%If trim(p_type) = "1" Or trim(p_type) = "" Then%> style="border: 3px solid red" <%Else%> style="border: 2px solid #C0C0C0" <%End If%> <%If is_send = false Then%> class=hand onClick=" chanFieldsObj(1,this)" <%End If%> name="FIELD">
	 </td></tr>
	 <tr>
	 <td name="FIELD" valign=top>
	 <img src="../../images/checkbox<%=img_%>.gif" <%If trim(p_type) = "5" Then%> style="border: 3px solid red" <%Else%> style="border: 2px solid #C0C0C0" <%End If%> <%If is_send = false Then%> class=hand onClick=" chanFieldsObj(5,this)" <%End If%> name="FIELD">
	 </td>
	 <td name="FIELD" valign=top>
	 <img src="../../images/radio<%=img_%>.gif"  <%If trim(p_type) = "8" Then%> style="border: 3px solid red" <%Else%>  style="border: 2px solid #C0C0C0" <%End If%> <%If is_send = false Then%> class=hand onClick=" chanFieldsObj(8,this)" <%End If%> name="FIELD">
	 </td>
	 <td name="FIELD" valign=top>
	 <img src="../../images/scale<%=img_%>.gif" <%If trim(p_type) = "9" Then%> style="border: 3px solid red" <%Else%> style="border: 2px solid #C0C0C0" <%End If%> <%If is_send = false Then%> class=hand onClick=" chanFieldsObj(9,this)" <%End If%> name="FIELD">
	 </td></tr>
	 <tr>
	 <td name="FIELD" valign=top>
	 <img src="../../images/date<%=img_%>.gif" <%If trim(p_type) = "6" Then%> style="border: 3px solid red" <%Else%> style="border: 2px solid #C0C0C0" <%End If%> <%If is_send = false Then%> class=hand onClick=" chanFieldsObj(6,this)" <%End If%> name="FIELD">
	 </td>
	 <td name="FIELD" valign=top>
	 <img src="../../images/radio1<%=img_%>.gif"  <%If trim(p_type) = "11" Then%> style="border: 3px solid red" <%Else%>  style="border: 2px solid #C0C0C0" <%End If%> <%If is_send = false Then%> class=hand onClick=" chanFieldsObj(11,this)" <%End If%> name="FIELD">
	 </td>
	 <td name="FIELD" valign=top>
	 <img src="../../images/scale1<%=img_%>.gif" <%If trim(p_type) = "12" Then%> style="border: 3px solid red" <%Else%> style="border: 2px solid #C0C0C0" <%End If%> <%If is_send = false Then%> class=hand onClick=" chanFieldsObj(12,this)" <%End If%> name="FIELD">
	 </td></tr>
	 </table>
	 <INPUT type="hidden" id=type_field name=type_field value="<%=p_type%>">
	</td>	
</tr>

<tr id=align_tr style="display:<%if showAlign = true then%>block<%else%>none<%end if%>">
	<td align="<%=align_var%>">
	 <table cellpadding=0 cellspacing=0 width=80 align="<%=align_var%>" border=0>
	 <tr>	 
	 <td width=40 align="<%=align_var%>">
	   <INPUT type=image name="DIR" id=word14 title="<%=arrTitles(14)%>" src="../../../htmlarea/images/dirltr.gif" value="ltr" <%if trim(p_align)="ltr" then%> style="border:3px solid red" <%Else%> style="border:2px solid #C0C0C0" <%end if%> onclick="return changeDir('ltr',this);return false;">
	 </td>
	 <td width=40 align="<%=align_var%>">
	   <INPUT type=image name="DIR" id=word15 title="<%=arrTitles(15)%>" src="../../../htmlarea/images/dirrtl.gif" value="rtl" <%if trim(p_align)="rtl" then%> style="border:3px solid red" <%Else%> style="border:2px solid #C0C0C0" <%end if%> onclick="return changeDir('rtl',this);return false;" ID="Image1">
	 </td>   
	 </tr>	 
	</table>
	 <input type=hidden name="type_align" id="type_align" value="<%=p_align%>">
	</td>
	<th class="title_show_form" align="<%=align_var%>"  nowrap><span id="word4" name=word4><!--כיוון השדה--><%=arrTitles(4)%></span>&nbsp;</th>
</tr>
<tr id=size_tr style="display:<%if showSize = true then%>block<%else%>none<%end if%>">
<td width=100% colspan=2>
<table cellpadding=0 cellspacing=0 width=100%>
<tr>
	<td width=100% align="<%=align_var%>"><input class="passw" name="size" size=5 value="<%=p_size%>" style=" TEXT-ALIGN: center"></td>
	<th class="title_show_form" align="<%=align_var%>" width=150 nowrap><span id="word5" name=word5><!--גודל שדה מקסימלי--><%=arrTitles(5)%></span>&nbsp;</th>
</tr>	
</table>	
</td>	
</tr>
<tr>
	<td width=100% align="<%=align_var%>"><input class="passw" name="must" type="checkbox"  <% If p_must = "True" Then %> checked <%End If%> ></td>
	<th class="title_show_form" align="<%=align_var%>" width=150 nowrap><span id="word6" name=word6><!--שדה חובה--><%=arrTitles(6)%></span>&nbsp;</th>
</tr>
<tr id=scale_tr style="display:<%if not IsNull(p_scale) then%>block<%else%>none<%end if%>">
	<td align="<%=align_var%>">
	<select  dir="ltr" name="selectscale" class="norm" >
		 <option value="2" <%if p_scale="2" then%> selected<%end if%>>2</option>
		 <option value="3" <%if p_scale="3" then%> selected<%end if%>>3</option>
		 <option value="4" <%if p_scale="4" or IsNull(p_scale) then%> selected<%end if%>>4</option>
		 <option value="5" <%if p_scale="5" then%> selected<%end if%>>5</option>
		 <option value="6" <%if p_scale="6" then%> selected<%end if%>>6</option>
		 <option value="7" <%if p_scale="7" then%> selected<%end if%>>7</option>
		 <option value="8" <%if p_scale="8" then%> selected<%end if%>>8</option>
		 <option value="9" <%if p_scale="9" then%> selected<%end if%>>9</option>
		 <option value="10" <%if p_scale="10" then%> selected<%end if%>>10</option>
	</select>
	</td>
	<th class="title_show_form" align="<%=align_var%>"  nowrap><span id="word7" name=word7><!--סקלה--><%=arrTitles(7)%></span>&nbsp;</th>
</tr>
<%if p_type="3" or p_type="8" or p_type="11" then %>
<tbody  id="options_tr" name="options_tr">
<tr><td align="<%=align_var%>" colspan=2 height=10 nowrap></td></tr>
<tr><td align="<%=align_var%>" colspan="2">
<table cellpadding=0 cellspacing=0 width=100% align="<%=align_var%>">
<tr>
	<td align="<%=align_var%>" width=100%><INPUT class="but_menu" type="button" value="<%=arrButtons(5)%>" style="width:70" onclick="JavaScript:visFormElem('','0')" id=Button5 name=Button5></td>
	<td align="<%=align_var%>" width=150 nowrap class="title_table_admin"><span id="word8" name=word8><!--אפשרויות בחירה--><%=arrTitles(8)%></span>&nbsp;</td>
</tr></table></td>	
</tr>
<%  Set p_select=con.GetRecordSet("Select * from Form_Select where Field_id="& id &" Order by Field_Value")
    If not  p_select.eof then%>      
    <tr><td align="<%=align_var%>" colspan=2>
    <a name="options_table"></a>
	<table width="490" cellspacing="1" cellpadding="1" align="<%=align_var%>">
	<tr>
		<td class="title_sort" align="center" width="50" nowrap>&nbsp;<span id="word9" name=word9><!--מחק--><%=arrTitles(9)%></span>&nbsp;</td>
		<td class="title_sort" align="center" width="50" nowrap>&nbsp;<span id="word10" name=word10><!--עדכן--><%=arrTitles(10)%></span>&nbsp;</td>
		<td class="title_sort" align="<%=align_var%>" width="100%">&nbsp;&nbsp;</td>
	</tr>
		<%do while not p_select.eof
			prText=p_select("Field_Value")
			prId=p_select("ID")
		%>
		<tr>
			<td align="center" class="card"><a href="addfield.asp?IdSelect=<%=p_select("ID")%>&selectdel=1&prodId=<%=prodId%>&id=<%=id%>&update=1" ONCLICK="return CheckDel()"><IMG SRC="../../images/delete_icon.gif" BORDER=0 name=word16 title="<%=arrTitles(16)%>"></a>&nbsp;</td>
			<td align="center" class="card"><a href="JavaScript: visFormElem('<%=altFix(prText)%>','<%=prId%>');"><IMG SRC="../../images/edit_icon.gif" BORDER=0 name=word17 title="<%=arrTitles(17)%>"></a>&nbsp;</td>
			<td align="<%=align_var%>" class="card" dir="<%=dir_obj_var%>">&nbsp;<%=p_select("Field_Value")%></td>
		</tr>
	    <%p_select.MoveNext 
            loop%>
  	</table>
	</td>	
	</tr>	
	<tr><td align="<%=align_var%>" colspan=2 height=5 nowrap></td></tr> 
	</tbody> 
 <% end if 'if not  p_select.eof then%>
<%End If%>
<tr><td align="<%=align_var%>" colspan="2" height=10 nowrap></td></tr>
<tr><td align="<%=align_var%>" colspan="2">
<table cellpadding=0 cellspacing=0 width=450 align="<%=align_var%>" border=0>
<tr><td width=45% align="center" nowrap>
<INPUT class="but_menu" style="width:90" type="button" value="<%=arrButtons(2)%>" id=button2 name=button2 onclick="document.location.href='addform.asp?prodId=<%=prodid%>#link<%=prodID%>'"></td>
<td width=50 nowrap></td>
<td width=45% align=center nowrap>
<input type=submit value="<%=arrButtons(1)%>" class="but_menu" style="width:90" onClick="return CheckFields()" id=button1 name=button1>
		<input type="hidden" name="new_new" value="true">
		<input type="hidden" name="prodId" value="<%=prodId%>">
		<input type="hidden" name="Id" value="<%=Id%>">
		<input type="hidden" name="update" value="true">		
	</td>
</tr>
</table></td></tr>
</form>
</table>
<td width=110 nowrap align="<%=align_var%>" valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width=100%>
<tr><td align="<%=align_var%>" colspan=2 height="17" nowrap></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;" href='addform.asp?prodId=<%=prodId%>'><span id=word11 name=word11><!--עריכת טופס--><%=arrTitles(11)%></span></a></td></tr>
</table></td></tr>
<tr><td height=10 nowrap></td></tr>
</table>

<div id="divText" name="divText">
<table border="0" cellspacing="1" cellpadding="0" align="center" width="460" bgcolor="#888888">
      <tr>
<!-- *** TEXT *** TEXT *** TEXT ***  -->
        <td width="100%" valign="top" align="<%=align_var%>">
         <table border="0" width="100%" cellspacing="0" cellpadding="0">
          <tr>
            <td width="100%">
            <table border="0" width="100%" cellspacing="0" cellpadding="0">
            <tr>
                   <td width="100%" colspan="3" class="title_form" align="<%=align_var%>">
                      <table border="0" width="100%" cellspacing="0" cellpadding="0">    
                         <tr>                                                        
                             <td align="center"><p id=titleEd name=titleEd style="font-size:10pt;color:#ffffff">&nbsp;</p></td>                             
                          </tr>
                      </table>
                   </td>  
            </tr>
            <tr>
                 <td width="100%" bgcolor="#E5E5E5"> 
                    <table border="0" width="100%" cellspacing="0" cellpadding="0">
                        <tr><td  height="5"></td></tr>
                        <form name="form1" action="addfield.asp?id=<%=id%>" method="post" onSubmit="return onSubmit_Text();" >  
                        <tr>
                         <td align="<%=align_var%>" width="100%">
                           <input type="hidden" name="idInf" value="" ID="idInf">
                           <input type="hidden" name="prodId" value="<%=trim(prodId)%>" ID="prodId">
                           <table border="0" width="100%" align="center" cellspacing="0" cellpadding="0">
								<tr>
									<td width=90% align="<%=align_var%>"><INPUT dir="<%=dir_obj_var%>" class="passw" type="text" id=textItem name=textItem  size=70 maxlength=100></td>
									<th class="title_show_form" width=10% align="<%=align_var%>">&nbsp;&nbsp;</th>
								</tr>
                            </table>
                           </td>
                         </tr>
                         <tr><td  align="center" height="35">
                         <table border="0" width="100%" align="center" cellspacing="0" cellpadding="0" dir="<%=dir_obj_var%>">
						 <tr><td width=50% align="center">
                         <A class="but_menu" href="#" onClick="return onSubmit_Text();" style="width:70" id=word12 name=word12><!--אישור--><%=arrTitles(12)%></a></td>
                         <td width=50% align="center">
                         <A class="but_menu" href="#" onClick="return hideFormElem();" style="width:70" id=word13 name=word13><!--ביטול--><%=arrTitles(13)%></a>
                         </td></tr></table></td></tr>
                         <tr><td  height="5"></td></tr>                                               
                        </form>  
                      </table>
                   </td>
              </tr>
            </table>
            </td>
          </tr>
        </table>
</td>
</tr>
<%END IF%>
</table>
</center></div></table>
<%If Request.QueryString("viewPopup") <> nil Then%>
<SCRIPT LANGUAGE=javascript>
<!--
visFormElem('',0);
//-->
</SCRIPT>
<%End If%>
<script language=javascript>
<!--
	 if(window.document.all('Title'))
		window.document.all('Title').focus();
//-->
</script>
</body>
<%set con=nothing%>
</html>




