<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%
  contactID = trim(Request("contactID"))
  companyID = trim(Request("companyID"))
  lang_id = trim(Request.Cookies("bizpegasus")("LANGID"))
  If lang_id = "2" Then
		dir_var = "rtl"
		align_var = "left"
		dir_obj_var = "ltr"
  Else
		dir_var = "ltr"
		align_var = "right"
		dir_obj_var = "rtl"
  End If		
	  
  letter = trim(Request.Form("letter"))
  fieldID = trim(Request("fieldID"))
  If trim(lang_id) = "1" Then
	Response.CharSet = "windows-1255"
	Session.LCID = 1037
  Else
    Response.CharSet = "windows-1252"
    Session.LCID = 2057
  End If	
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<style>
A.button_letter
{
    BORDER-RIGHT: #4A4978 2px solid;
    BORDER-BOTTOM: #4A4978 2px solid;
    BORDER-TOP: #B3B2D0 2px solid;
    BORDER-LEFT: #B3B2D0 2px solid;
    FONT-WEIGHT: bold;
    FONT-SIZE: 9pt;
    COLOR: #FFFFFF;
    FONT-FAMILY: Arial;
    HEIGHT: 22px;
    width: 22px;
    BACKGROUND-COLOR: #6F6DA6;
    TEXT-DECORATION: none;
    text-align: center;
}
A.button_letter:hover
{  
    COLOR: #E9C234;
}

A.button_letter_active
{
    BORDER-RIGHT: #B3B2D0 2px solid;
    BORDER-BOTTOM: #B3B2D0 2px solid;
    BORDER-TOP: #4A4978 2px solid;
    BORDER-LEFT: #4A4978 2px solid;
    HEIGHT: 22px;
    width: 22px;
    FONT-WEIGHT: bold;
    FONT-SIZE: 9pt;
    COLOR: #FFFFFF;
    FONT-FAMILY: Arial;
    BACKGROUND-COLOR: #E9C234;
    TEXT-DECORATION: none;
    text-align: center
 }
</style>
<script>
	var selectAscii=0;
	function SubForm(selectAscii)
	{
		if (selectAscii == 0)
			window.form1.letter.value='';
		window.form1.submit();
		return true;
	}
	
	//!!!!!!!MILA - New design - text buttons
	function FormLetterValue(letterValue,ascii)
	{
		if (selectAscii != 0)
			changeclass("key"+selectAscii,"button_letter");
		changeclass("key"+ascii,"button_letter_active");
		selectAscii=ascii;
		document.forms["form1"].letter.value=letterValue;
		SubForm(selectAscii);
	}
	
	function goBack(cityID,cityObj)
	{
		cityName = cityObj.innerText;
		fieldID = "<%=trim(fieldID)%>";
		if(fieldID == "")
			fieldID = "cityName";		
		if(window.opener.document.all(fieldID))
		{
			window.opener.document.all(fieldID).value = cityName;
		}
		if(window.opener.document.all("cityID"))
		{
			window.opener.document.all("cityID").value = cityID;
		}
		window.opener.focus();
		self.close();		
	}
</script>

<script language="VBScript">
	function changeclass(id,clname)
		document.all(id).classname=clname
	end function
</script>
</head>

<body style="margin:0;background:'#e6e6e6'">
<table  border="0" cellpadding="0" cellspacing="0" width="100%">		
<tr><td class="page_title"><%If trim(lang_id) = "1" Then%>רשימת הערים<%Else%>Countries list<%End If%></td></tr>
<form action="cities_list.asp" id="form1" name="form1" method=post>
<tr>
	<td width="100%">
	<input type="hidden" name="letter" ID="letter" value="">
	<input type="hidden" name="fieldID" ID="fieldID" value="<%=fieldID%>">	
	<table width="90%" width="0" border="0" cellpadding="2" cellspacing="5" align=center>
		<tr>
		<%If trim(lang_id) = "1" Then%>
		<%for ascii=235 to 224 step -1%>
			<%if ascii<>234 then%>
			<td>		
			<a class="button_letter<%If trim(letter) = trim(Chr(ascii)) Then%>_active<%End If%>" id="key<%=ascii%>" name="key<%=ascii%>" href="javascript:FormLetterValue('<%=Chr(ascii)%>','<%=ascii%>')">&nbsp;<%=Chr(ascii)%>&nbsp;</a>
			</td>
			<%end if%>
		<%next%>
		<%Else%>		
		<%for ascii=65 to 77 step 1%>					
			<td>		
			<a class="button_letter<%If trim(letter) = trim(Chr(ascii)) Then%>_active<%End If%>" id="key<%=ascii%>" name="key<%=ascii%>" href="javascript:FormLetterValue('<%=Chr(ascii)%>','<%=ascii%>')">&nbsp;<%=Chr(ascii)%>&nbsp;</a>
			</td>		
		<%next%>
		
		<%End If%>
		</tr>
		
		<tr>
		<%If trim(lang_id) = "1" Then%>
			<%for ascii=250 to 236 step -1%>
			<%if ascii<>237 and ascii<>239 and ascii<>243 and ascii<>245 then%>
			<td>	
			<a class="button_letter<%If trim(letter) = trim(Chr(ascii)) Then%>_active<%End If%>" name="key<%=ascii%>" name="key<%=ascii%>" href="javascript:FormLetterValue('<%=Chr(ascii)%>','<%=ascii%>')">&nbsp;<%=Chr(ascii)%>&nbsp;</a>
			</td>	
			<%end if%>
		<%next%>
		<%Else%>
			<%for ascii=78  to 90 step 1%>	
			<td>	
			<a class="button_letter<%If trim(letter) = trim(Chr(ascii)) Then%>_active<%End If%>" name="key<%=ascii%>" name="key<%=ascii%>" href="javascript:FormLetterValue('<%=Chr(ascii)%>','<%=ascii%>')">&nbsp;<%=Chr(ascii)%>&nbsp;</a>
			</td>
			<%next%>			
		<%End If%>		
		</tr>
	</table></td></tr>
	</form>
	<%
	If Request.Form("letter") <> nil Then
	    numOfCities = 0
	    If trim(lang_id) = "1" Then	   
		sqlstr = "Select city_id, city_name From cities WHERE ltrim(rtrim(city_name)) Like '" & sFix(letter) & "%' Order BY city_name"
		Else
		sqlstr = "Select city_id, city_Name_Eng From cities WHERE ltrim(rtrim(Lower(city_Name_Eng))) Like '" & sFix(LCase(letter)) & "%' Order BY city_Name_Eng"
		End If
		set rs_cities = con.getRecordSet(sqlstr)
		if not rs_cities.eof Then
			numOfCities = rs_cities.recordCount
			cities_arr = rs_cities.getRows()
			
		end if
		set rs_cities = Nothing
		'Response.Write numOfCities
		'Response.End
	    If numOfCities > 0 Then	
	%>
	<tr>
	<td width="100%">	
	<table width="90%" width="0" border="0" cellpadding="2" cellspacing="0" align=center dir="<%=dir_obj_var%>">	
	<%If numOfCities > 20 Then%>
	<tr><td width=50% nowrap valign=top><table cellpadding=0 cellspacing=0 width=100%>
	<%For i=0 To Fix((numOfCities/2)+0.9)-1%>
	<tr><td align="<%=align_var%>">&nbsp;<a href="" onclick="goBack('<%=cities_arr(0,i)%>',this)" class="link1"><%=cities_arr(1,i)%></a>&nbsp;</td></tr>
	<%Next%>
	</table></td>
	<td width=50% nowrap valign=top><table cellpadding=0 cellspacing=0 width=100% border=0>
	<%For i=Fix((numOfCities/2)+0.9) To numOfCities-1%>
	<tr><td align="<%=align_var%>">&nbsp;<a href="" onclick="goBack('<%=cities_arr(0,i)%>',this)" class="link1"><%=cities_arr(1,i)%></a>&nbsp;</td></tr>
	<%Next%>
	<%Else%>
	<%For i=0 To numOfCities-1%>
	<tr><td align="<%=align_var%>">&nbsp;<a href="" onclick="goBack('<%=cities_arr(0,i)%>',this)" class="link1"><%=cities_arr(1,i)%></a>&nbsp;</td></tr>
	<%Next%>
	<%End If%>
	</td></table>
	</tr>
	<%	End If %>			
	</table></td></tr>	
	<%	End If %>	
	</table>		
</body>
</html>
<%
set con = nothing
%>
