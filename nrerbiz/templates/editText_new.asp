<%SERVER.ScriptTimeout=3000%>
<!--#INCLUDE file="..\..\netcom/reverse.asp"-->
<!--#include file="..\..\netcom/connect.asp"-->
<!--#include file="../checkWorker.asp"-->
<%

   elemId=trim(Request("elemId"))
   pageId=trim(Request("pageId"))
   prodId=trim(Request("prodId"))
   place=CInt(Request("place"))
   innerparent=request("innerparent")

   newPlace=place+1
   persize=2
  'perHebrSize=100
  If isNumeric(elemId) = true then
  selm="select * from Template_Elements where Element_Id="&elemId&" "
  set pr=con.getRecordSet(selm)
   pageId=pr("Page_Id")
   pertext=pr("Element_Text")
   perLink=pr("Element_Link")
   linkStatus=Trim(pr("Element_Link_Status"))
   linkTarget=Trim(pr("Element_Link_Target"))
   pBullet=pr("Element_Bullet")
   pWord=pr("Element_Word")
   wPercolor=pr("Elemen_Word_Color")
   wPercolorname=pr("Element_Word_Color_Name") 
   wPersize=pr("Element_Word_Size")
   wPertype=pr("Element_Word_Type") 
   persize=pr("Element_Font_Size")
   pertype=pr("Element_Font_Type")
   percolor=pr("Element_Font_Color")
   percolorname=pr("Element_Font_Color_Name")
   perbgcolor=pr("Element_BGColor")
   perbgcolorname=pr("Element_BGColor_Name")
   perHebrSize=pr("Element_Hebrew_Size")
   peralign=pr("Element_Align")
   perSpace=pr("Element_Space")
   pr.close 
  end if
  set pr = Nothing
    
 if pBullet=nil then
	pBullet="0" 
 end if  
 if wPercolor=nil then
	wPercolor="#000000"
 end if  
 if wPercolorname=nil then
	wPercolorname="Black"
 end if  
 if wPersize=nil then
	wPersize=2
 end if  
 if wPertype=nil then
	wPertype="Regular"
 end if  
 if persize=nil then
	persize=2
 end if  
 if perHebrSize=nil then
	perHebrSize=85
 end if
 if linkStatus=nil then
	linkStatus="noLink"
 end if  
 if percolorname=nil then
	percolorname="Black"
 end if  
 if percolor=nil then
	percolor="#000000"
 end if  
 if peralign=nil then
	peralign="right"
 end if
%>
<html>
<head>
<title>ניהול תבניות דפי מבצע</title>
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<link href="../forms.css" rel="STYLESHEET" type="text/css">

<script language="JavaScript">
<!--                                    
function mywindow(){
    str="color.htm";
  
	numsave=open(str,"","scrollbars=yes,menubar=no,width=300,height=500");
	numsave.document.close();
	
	return false;
}
function mywindow1(){
    str="color1.htm";
  
	numsave=open(str,"","scrollbars=yes,menubar=no,width=300,height=500");
	numsave.document.close();
	
	return false;
}
function mywindow2(){
    str="color2.htm";
  
	numsave=open(str,"","scrollbars=yes,menubar=no,width=300,height=500");
	numsave.document.close();
	
	return false;
}

function CheckFields(){
	  if (document.f1.title1.value=='')
		{
		alert('חובה למלא את השדה טקסט');
		return false;
		}
	  else
		return true;
}

function FormSubmitFun(){
	if (CheckFields())
	document.f1.submit();
return false;
}
function funclink(off1,off2,on)
{
    window.document.forms[0].elements[off1].checked=false;
	window.document.forms[0].elements[off2].checked=false;
	window.document.forms[0].elements[on].checked=true;
	
/*    window.document.forms[0].Template_Elements[on].checked=false;
	window.document.forms[0].Template_Elements[off2].checked=false;
	window.document.forms[0].Template_Elements[off1].checked=true;*/
}
//-->
</script>
</head>
<body topmargin=0 leftmargin=0 rightmargin=0 bottommargin=0>
<table bgcolor="#c0c0c0" border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="5"  bgcolor="#ff9100"  align="center" valign="middle" nowrap><font size="4" color="#ffffff"><strong><B>עדכון טקסט מבצע</B>&nbsp;</strong></font></td>
  </tr> 
</table>
<table border="0" width="100%"  cellspacing="1" cellpadding="0">
	<tr>
		<td align="center"colspan=2 width=100%>&nbsp;</td>
			</tr>
</table>
<table width="710" cellspacing="0" cellpadding="0" align=center border="0" bgcolor="#ffffff" >

<tr><td>
<table border="0" width="100%" align="center" cellspacing="1" cellpadding="0" bgcolor="#FFFFFF">
  <tr>
    <td align="center" valign="middle" bgcolor="#3E596E" nowrap class="small_table_header">טקסט</td>
  </tr>
<tr>
  <td bgcolor="#DDDDDD" height="10"><table><tr><td></td></tr></table></td>
</tr>
<form name="f1" id="f1" ACTION="event.asp?innerparent=<%=innerparent%>&prodId=<%=prodId%>&pageId=<%=pageId%>" METHOD="post">
<tr>
  <td align="center" bgcolor="#DDDDDD">
	 <font face="Arial (Hebrew)"><textarea name="title1" dir="rtl" rows="8" cols="80"><%=pertext%></textarea></font>
  </td>
</tr>
<tr>
  <td bgcolor="#DDDDDD" height="10" nowrap></td>
</tr>
 <tr><td >
 <table width=100% bgcolor="#DDDDDD" border="0" cellspacing="0" cellpadding="0"> 
 <%if (linkStatus<>"noLink" and elemId<>nil) then
	set wrd=con.getRecordSet("SELECT * FROM Link_Texts where Element_Id="&elemId&" ")
		if not wrd.eof then
			wordLink=wrd("Key_Text")
			lnk=wrd("Link_Text")
			clr=wrd("Link_Color")
		end if
	wrd.close
  end if	
	%>
  <tr>
  	<td width="50%" nowrap valign="top">
  	<table border=0 cellpadding=0 cellspacing=0 width=100%>
  	<tr><td colspan="4" align="center" bgcolor="#3E596E" class="small_table_header">קישור</td></tr>
  	<tr><td colspan="4" nowrap>
  	<table border=0 cellpadding=0 cellspacing=0 width=100%>
  	<td width=33% bgcolor="#DDDDDD" align="center" class="admin_form_title">
		<input type="radio" name="linkTarget" value="_blank" <%if linkTarget="_blank" then%>checked<%end if%> onclick="funclink(3,2,1)">
		קישור חיצוני
	</td>	
  	<td width=33% bgcolor="#DDDDDD" align="center" class="admin_form_title">
		<input type="radio" name="linkTarget" value="_self" <%if linkTarget="_self" then%>checked<%end if%> onclick="funclink(3,1,2)">
	    קישור פנימי
	</td>	
  	<td width=33% bgcolor="#DDDDDD" align="center" class="admin_form_title">
		<input type="radio" name="isLink" value="noLink" <%if linkStatus="noLink" then%>checked<%end if%> onclick="funclink(2,1,3)">
		ללא קישור
	</td>
	</tr>
	</table></td></tr>
	<tr>
    <td bgcolor="#DDDDDD" width="50%" align="left" nowrap><input type="text" class="texts" name="allLink" size="30" <%if linkStatus="global" then%> value="<%=perLink%>"<%end if%>></td>
    <td bgcolor="#DDDDDD" width="17%" align="left" class="form" nowrap>קישור ל</td>
	<td bgcolor="#DDDDDD" width="10%" align="right" nowrap><input type="radio" name="isLink" value="global" <%if linkStatus="global" then%>checked<%end if%>></td>
	<td bgcolor="#DDDDDD" width="23%" align="left"  class="admin_form_title" nowrap>&nbsp;כל הטקסט</td>	
  </tr>  
   <tr>
    <td bgcolor="#DDDDDD" width="50%" align="left" nowrap><font face="Arial (Hebrew)"><input dir="rtl" type="text" class="texts" name="word1" size="30" <%if linkStatus="words" then%>value="<%=wordLink%>" <%end if%>></font></td>
    <td bgcolor="#DDDDDD" width="17%" align="left" class="form" nowrap>&nbsp;:מילה</td>
	<td bgcolor="#DDDDDD" width="10%" align="right" nowrap><input type="radio" name="isLink" value="words" <%if linkStatus="words" then%>checked<%end if%>></td>
	<td bgcolor="#DDDDDD" width="23%" align="left" class="admin_form_title" nowrap>&nbsp;מילה</td>
  </tr>
  <tr>
    <td bgcolor="#DDDDDD" width="50%" align="left" nowrap><input type="text" class="texts" name="wordLink1" size="30" <%if linkStatus="words" then%>value="<%=lnk%>"<%end if%>></td>
    <td bgcolor="#DDDDDD" width="17%" align="left" class="form" nowrap>קישור ל</td>
	<td bgcolor="#DDDDDD" colspan="2" width=33% nowrap>&nbsp;</td>
  </tr>
</table>	
	</td>
	<td width="1" nowrap bgcolor="#FFFFFF"></td>
	<td width="20%" bgcolor="#DDDDDD" valign="top" nowrap>
		<table bgcolor="#DDDDDD" border="0" cellspacing="0" cellpadding="0" width="100%">
			<tr><td align="center" bgcolor="#3E596E" class="small_table_header">יישור</td></tr>
			<tr>
				<td align="center" bgcolor="#DDDDDD" nowrap>
				    <font face="Arial (Hebrew)"><div dir="rtl">
				  	<select class="app" style="width:80" name="falign" size="3"> 
					    <option value="center" <%if peralign="center" then%>selected<%end if%>><%="ממורכז"%></option>
						<option value="left" <%if peralign="left" then%>selected<%end if%>><%="לשמאל"%></option>
    					<option value="right" <%if peralign="right" then%>selected<%end if%>><%="לימין"%></option>
					</select>
					</div></font>
				</td>
			</tr>
			<tr><td bgcolor="#DDDDDD" height="5" nowrap></td></tr>
			<tr>
				<td align="center" bgcolor="#DDDDDD" class="form" nowrap>
				<%i=0%>
				<select class="app" style="width:50" name="tSpace" size="1">
					<%do while i<=200%>
					<option value="<%=i%>" <%if perSpace=i then%>selected<%end if%>><%=i%></option>
				<%i=i+5
				  Loop%>
				</select>:הצקהמ חוור
				</td>
			</tr>
		</table>	
	</td>
	<td width="1" bgcolor="#FFFFFF" nowrap></td>
	<td width="30%" bgcolor="#DDDDDD" valign="top" nowrap>
		<table bgcolor="#DDDDDD" border="0" cellspacing="0" cellpadding="0" width="100%">
			<tr><td colspan="2" align="center" bgcolor="#3E596E" class="small_table_header">גופן</td></tr>
			<tr>
				<td bgcolor="#DDDDDD" align="right" nowrap>
				<input type="text" class="texts" name="fcolnm" style="width:95%" value="<%=percolorname%>">
					<input type="hidden" name="fcolor" value="<%=percolor%>">
					<input type="hidden" name="fcolname" value="<%=percolorname%>">
				</td>
				
				<td align="left" bgcolor="#DDDDDD" nowrap>&nbsp;<A  class="linkFaq" HREF="" ONCLICK="return mywindow();"><%="צבע טקסט"%></a></td>
				
				<!--
				<td >
				    <table  border="0" cellspacing="0" cellpadding="0" width="100%">
 			          <tr>
				      <td width="16" nowrap name="colorSelect"  id=colorSelect title="בחר צבע" class=coolButton onclick="paletteToggle()">
		                 <IMG align=absMiddle name=colorSelect src="../images/paletteicon.gif">
	                  </td>
	                  <td width="100%"></td>
	                  <td style="WIDTH:0px;" name="colorSelect" >
	                    <OBJECT id=palette style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 245px" type=text/x-scriptlet data="" VIEWASTEXT></OBJECT>
		                <script	type="text/javascript" for="palette" event="onscriptletevent(name, eventData)">
			               switch(name){
				           case "colorchange":
					          doFormat("ForeColor",eventData);
					          paletteToggle();
			               }
		                </script>
	                  </td>
	                 </tr>
	              </table> </td>
	              -->
	             </tr>

	                
	              </table>
	            </td>
			</tr>

			<tr>
				<td bgcolor="#DDDDDD" align="right" nowrap>
				<input type="text" class="texts" name="bgcolnm" style="width:95%" value="<%=perbgcolorname%>">
					<input type="hidden" name="bgcolor" value="<%=perbgcolor%>">
					<input type="hidden" name="bgcolname" value="<%=perbgcolorname%>">
				</td>
				<td align="left" bgcolor="#DDDDDD" nowrap>&nbsp;<A  class="linkFaq" HREF="" ONCLICK="return mywindow2();">צבע רקע</a></td>
			</tr>
			<tr>
				<td bgcolor="#DDDDDD" align="right" nowrap>
				   <select class="app" style="width:80" name="fsize" size="1"> 
					<%i=1
					  do while i<=7 %>
    <option value="<%=i%>" <%if persize=i then%>selected<%end if%>><%=i%></option>
					<%i=i+1
					  Loop%>
				   </select>
				</td>
				<td align="left" bgcolor="#DDDDDD" class="form">&nbsp;:גודל</td>
			</tr>
			<tr>
				<td bgcolor="#DDDDDD" align="right" nowrap>
   					<font face="Arial (Hebrew)">
					<select class="app" style="width:80" name="fstyle" size="1"> 
						<option value="Regular">רגיל</option>
						<option value="I" <%if pertype="I" then%>selected<%end if%>>נטוי</option>
						<option value="STRONG" <%if pertype="STRONG" then%>selected<%end if%>>מודגש</option>
						<option value="U" <%if pertype="U" then%>selected<%end if%>>קו תחתון</option>
						<option value="SUB" <%if pertype="SUB" then%>selected<%end if%>>כתב עילי</option>
						<option value="SUP" <%if pertype="SUP" then%>selected<%end if%>>כתב תחתי</option>
						<option value="STRIKE" <%if pertype="STRIKE" then%>selected<%end if%>>קו חוצה</option>
					</select>
					</font>
				</td>
				<td align="left" bgcolor="#DDDDDD" class="form">&nbsp;:סגנון</td>
			</tr>
			<tr>
				<td bgcolor="#DDDDDD" align="right" nowrap>
					<%i=5%>
					<font face="Arial (Hebrew)">
					<select class="app" style="width:80" name="hebrSize" size="1">
						<option value="null">חופשי</option>
						<%do while i<=100%>
						<option value="<%=i%>" ><%=i%></option>
						<%i=i+1
					      loop%>
					</select> 
					</font>
				</td>
				<td align="left" bgcolor="#DDDDDD" class="form">&nbsp;:אורך השורה</td>
			</tr>
		</table>	
	</td>			
			</tr>
		</table>
	</td>
</tr>
<!--///////////////////////////////////////-->
<tr>
	<td width="100%" valign="top" bgcolor="#DDDDDD">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" bgcolor="#DDDDDD">
	  <tr>
	  <td width="50%">
		<table bgcolor="#DDDDDD" border="0" cellspacing="0" cellpadding="0" width="100%">
			<tr><td colspan="2" align="center" bgcolor="#3E596E" height="20" class="small_table_header">מילה  בתוך  הטקסט</td></tr>
			<tr>
				<td bgcolor="#DDDDDD" align="right" nowrap class="form">
				   <select class="app" style="width:80" name="wSize" size="1"> 
					<%i=1
					  do while i<=7 %>
					<option value="<%=i%>" <%if wPersize=i then%>selected<%end if%>><%=i%></option>
					<%i=i+1
					  Loop%>
				   </select>&nbsp;&nbsp;&nbsp;:גודל
				</td>
				<td width="60%" align="center" valign="middle" bgcolor="#DDDDDD" nowrap class="form"><input type="text" class="texts" style="width:120" name="WWord" value="<%=pWord%>">&nbsp;:מילה</td>
			</tr>
			<tr>
				<td bgcolor="#DDDDDD" align="right" nowrap class="form">
   					<font face="Arial (Hebrew)">
					<select class="app" style="width:80" name="wStyle" size="1"> 
						<option value="Regular">רגיל</option>
						<option value="I" <%if wPertype="I" then%>selected<%end if%>>נטוי</option>
						<option value="STRONG" <%if wPertype="STRONG" then%>selected<%end if%>>מודגש</option>
						<option value="U" <%if wPertype="U" then%>selected<%end if%>>קו תחתון</option>
						<option value="SUB" <%if wPertype="SUB" then%>selected<%end if%>>כתב עילי</option>
						<option value="SUP" <%if wPertype="SUP" then%>selected<%end if%>>כתב תחתי</option>
						<option value="STRIKE" <%if wPertype="STRIKE" then%>selected<%end if%>>קו חוצה</option>
					</select>
					</font>&nbsp;:סגנון
				</td>
				<td bgcolor="#DDDDDD" align="center" nowrap><input type="text" class="texts" name="wColnm" style="width:120" value="<%=wPercolorname%>">
					<input type="hidden" name="wColor" value="<%=wPercolor%>">
					<input type="hidden" name="wColname" value="<%=wPercolorname%>">
				&nbsp;<A  class="linkFaq" HREF="" ONCLICK="return mywindow1();">צבע</a>
				</td>
			</tr>
			<tr><td colspan="2" bgcolor="#DDDDDD" width="100%" height="5"><table><tr><td></td></tr></table></td></tr>
		</table>
	</td>
	<td width="1" nowrap bgcolor="#FFFFFF"></td>
	<td width="50%" valign="top">
		<table bgcolor="#DDDDDD" border="0" cellspacing="0" cellpadding="0" width="100%" height="100%">
			<tr><td colspan="4" align="center" bgcolor="#3E596E" height="20" class="small_table_header">תבליטים</td></tr>
			<tr>
				<td bgcolor="#DDDDDD" align="right" nowrap>
			<tr><td colspan="4" bgcolor="#DDDDDD" width="100%" height="5"><table><tr><td></td></tr></table></td></tr>
			<tr>
				<td width="25%" bgcolor="#DDDDDD" align="right" nowrap><img src="../../Images/bul3.gif" border="0">&nbsp;<input type="radio" name="bullet" value="3" <%if pBullet="3" then%>checked<%end if%>></td>
				<td width="25%" bgcolor="#DDDDDD" align="right" nowrap><img src="../../Images/bul2.gif" border="0">&nbsp;<input type="radio" name="bullet" value="2" <%if pBullet="2" then%>checked<%end if%>></td>
				<td width="25%" bgcolor="#DDDDDD" align="right" nowrap><img src="../../Images/bul1.gif" border="1" >&nbsp;<input type="radio" name="bullet" value="1" <%if pBullet="1" then%>checked<%end if%>></td>
				<td width="25%" bgcolor="#DDDDDD" align="right" valign="top" nowrap><font size="2">ללא תבליטים</font>&nbsp;<input type="radio" name="bullet" value="0" <%if isNull(pBullet) or pBullet="0" then%>checked<%end if%>></td>
			</tr>
			<tr>
				<td width="25%" bgcolor="#DDDDDD" align="right" nowrap><img src="../../Images/bul7.gif" border="0" >&nbsp;<input type="radio" name="bullet" value="7" <%if pBullet="7" then%>checked<%end if%>></td>
				<td width="25%" bgcolor="#DDDDDD" align="right" nowrap><img src="../../Images/bul6.gif" border="0" >&nbsp;<input type="radio" name="bullet" value="6" <%if pBullet="6" then%>checked<%end if%>></td>
				<td width="25%" bgcolor="#DDDDDD" align="right" nowrap><img src="../../Images/bul5.gif" border="0" >&nbsp;<input type="radio" name="bullet" value="5" <%if pBullet="5" then%>checked<%end if%>></td>
				<td width="25%" bgcolor="#DDDDDD" align="right" nowrap><img src="../../Images/bul4.gif" border="0" >&nbsp;<input type="radio" name="bullet" value="4" <%if pBullet="4" then%>checked<%end if%>></td>
			</tr>
		</table>
		</td>
	  </tr>
	  <tr><td height=10 nowrap></td><td height=10 width=1 bgcolor="#FFFFFF" nowrap></td><td height=10 nowrap></td></tr>
	  </table>
	</td>
</tr>
<!--//////////////////////////////////////////////-->
	    <input type="hidden" name="elemId" value="<%=elemId%>">
	    <input type="hidden" name="pageId" value="<%=pageId%>">
	    <input type="hidden" name="homePage" value="<%if perHomePage=True then%>&quot;1&quot;<%else%>&quot;0&quot;<%end if%>">
	    <input type="hidden" name="place" value="<%=place%>">
	    <input type="hidden" name="editTxt" value="1">
<tr>
	<td width="100%" nowrap bgcolor="#DDDDDD" height=30 valign=middle>
	<table width="100%" border="0" cellpadding="1" cellspacing="0" bgcolor="#FFFFFF">
	  <tr>
		<td bgcolor="#DDDDDD" width=45% nowrap align="right">
	    <INPUT class="but" style="width:90" type="button" onclick="document.location.href='event.asp?innerparent=<%=innerparent%>&pageId=<%=pageId%>&amp;prodId=<%=prodId%>&isright=<%=isRightMenu%>'" value="ביטול" id=button2 name=button2>
		</td>
		<td width=5% nowrap bgcolor="#DDDDDD"></td>
		<td bgcolor="#DDDDDD" width=45% nowrap align="left">	    
	    <INPUT class="but" style="width:90" type="button" onclick="return FormSubmitFun();" value="אישור" id=button1 name=button1>
		</td>		
	  </tr>
	</table>
	</td>
</tr>	    
</form>
</table>
</td></tr>

</table>
<br>
</body>
</html>