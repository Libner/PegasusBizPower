<%SERVER.ScriptTimeout=3000%>
<!--#INCLUDE file="..\..\netcom/reverse.asp"-->
<!--#include file="..\..\netcom/connect.asp"-->
<!--#include file="../checkWorker.asp"-->
<html>
<head>
<title>ניהול תבניות דפי מבצע</title>
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<link href="../../admin.css" rel="STYLESHEET" type="text/css">
<script language="JavaScript">
<!--
function ifFieldEmpty(){
	  if (document.form1.UploadFile1.value=='')
		{
		alert('חובה לבחור את התמונה');
		return false;
		}
	  else
		return true;
}

function CheckFields(){
	  if (document.f1.stam.value<=0)
		{
		alert('אנא לבחור את התמונה');
		return false;
		}
	  else
		return true;
}

function FormSubmitFun(){
//	if (CheckFields())
//if (confirm("?האם ברצונך לשמור את השינוים"))
//{
	if (document.f1.numcolumn.checked) 
	{	get_mytable();	}
	document.f1.submit();
//}	
return false;
}

var winColor, winColor_1, winColor_2;
function mywindow(){
    str="color.htm";
    left = parseInt(window.screen.width/2 -120)
	winColor=open(str,"winColor","scrollbars=yes,menubar=no,width=120,height=350,left="+left+",top=10");
	winColor.document.close();
	
	return false;
}
function mywindow1(){
    str="color1.htm";
    left = parseInt(window.screen.width/2 -120)
	winColor_1=open(str,"winColor_1","scrollbars=yes,menubar=no,width=120,height=350,left="+left+",top=10");
	winColor_1.document.close();
	
	return false;
}
function mywindow2(){
    str="color2.htm";
    left = parseInt(window.screen.width/2 -120)
	winColor_2=open(str,"winColor_2","scrollbars=yes,menubar=no,width=120,height=350,left="+left+",top=10");
	winColor_2.document.close();
	return false;
}

function GetNumbers()
{
	var ch=event.keyCode;
	event.returnValue =(ch >= 48 && ch <= 57) || ch == 46 ;
}

//-->
</script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
	var html_text = new String();
	var row_number = new Number();
	var column_number = new Number();
	var border_number;
	var tbl_align;
	var tbl_width;
	
function clearHTMLfunc()
{
	if (document.all("clearHTML",[1]).checked){		// common text
		document.all("title1").disabled=false;
		document.all("title1").dir='rtl';
		document.all("numcolumn").checked=false;
		document.all("textpict").style.display='none';
		document.all("title1").value='';
		document.all("title1").cols=50;
		//document.all("propText1").style.display='block';
		//document.all("propText2").style.display='block';
		}
	else{											//HTML
		document.all("title1").value=titleText;
		document.all("title1").dir='ltr';
		document.all("title1").cols=75;
		//document.all("propText1").style.display='none';
		//document.all("propText2").style.display='none';
		}
}	
	
function button1_onclick() {
	row_number = document.f1.numrow.value;
	column_number = document.f1.columnnumb.value;
	border_number = document.f1.numborder.value;
	tbl_align = document.f1.tbl_align.value;
	tbl_width = document.f1.numwidth.value;
	html_text = "";
	html_text = "<TABLE BORDER="+ border_number +" width="+ tbl_width +" align="+ tbl_align +" CELLSPACING=1 CELLPADDING=1>";
	for (i=0;i<row_number;i++)
	{
		html_text = html_text + "<tr>";
		for (j=0;j<column_number;j++)
		{
		html_text = html_text + "<td>";
		html_text = html_text + "<textarea dir=rtl name=txtarea" + i + "_" + j + " rows=3 style='width:95%'></textarea>";
		html_text = html_text + "</td>";
		}
		html_text = html_text + "</tr>"	;
	}
	html_text = html_text + "</TABLE>";
	mytable.innerHTML = html_text;
	if (document.f1.txtarea0_0)
	{	document.f1.txtarea0_0.scrollIntoView();}
	//alert(html_text);
}

function numcolumn_onclick(pr) {

	if (document.all("numcolumn").checked==true){
		document.all("mytable").style.display='block';
		document.all("clearHTML_td").style.display='none';
		document.all("textPict").style.display='block';
		//document.all("propText1").style.display='none';
		//document.all("propText2").style.display='none';
		document.all("pictureHTML").value='1';
		document.all("text_near_image").style.display='none';
		}
	else{
		document.all("text_near_image").style.display='inline';
		document.all("textPict").style.display='none';
		document.all("clearHTML_td").style.display='block';
		document.all("mytable").style.display='none';
//alert(pr)
		if (pr!='<'){
			//document.all("propText1").style.display='block';
			//document.all("propText2").style.display='block';
			}
		else{
			document.all("clearHTML").checked=false;}
		document.all("pictureHTML").value='0'
		}	
/*	if (document.f1.numcolumn.checked)
	{mytable.innerHTML = "";}
	else
	{button1_onclick();}*/
}

function get_mytable() {

	widCol=100/column_number;
	html_text = "";
	html_text = "<TABLE BORDER="+ border_number +" width="+ tbl_width +" align="+ tbl_align +" CELLSPACING=1 CELLPADDING=1>";
	for (i=0;i<row_number;i++)
	{
		html_text = html_text + "<tr>";
		for (j=0;j<column_number;j++)
		{
			html_text = html_text + "<td align=right valign=top width="+widCol+"%>";
			if (document.f1.elements['txtarea'+ i + '_' + j].value == '')
			{
				html_text = html_text + "<font style='FONT-SIZE:10pt; font-family:Ariel(Hebrew)'><p>&nbsp;</p></font>";
			}	
			else
			{
				html_text = html_text + "<p>" + document.f1.elements['txtarea'+ i + '_' + j].value + "</p>";
			}	
			html_text = html_text + "</td>";
		}
		html_text = html_text + "</tr>"	;
	}
	html_text = html_text + "</TABLE>";
	document.f1.htmltext2.value = html_text;
	
	//alert(html_text);
}
//-->
</SCRIPT>
</head>

<%elemId=Request("elemId")
  pageId=Request("pageId")
  prodId=Request("prodId")

  place=CInt(Request("place"))
  newPlace=place+1
  persize=2 
  perHebrSize=100
	if elemId<>nil then
	  set para=con.getRecordSet("select * from Template_Elements where Element_Id="&elemId&" ")
	    pageId=para("Page_Id")
	    pertext=para("Element_Text")
		perLinkStatus=Trim(para("Element_Link_Status"))
		perLink=para("Element_Link")
		linkTarget=Trim(para("Element_Link_Target"))
		perpicture=para.Fields("Element_Picture").ActualSize
		persize=para("Element_Font_Size")
		pertype=para("Element_Font_Type")
		percolor=para("Element_Font_Color")
		percolorname=para("Element_Font_Color_Name")
		perHebrSize=para("Element_Hebrew_Size")
		peralign=para("Element_Align")
		perTextImgAlign=para("Element_Text_Image_Align")
		perSpace=para("Element_Space")
		perbgcolor=para("Element_BGColor")
		perbgcolorname=para("Element_BGColor_Name")
		pBullet=para("Element_Bullet")
		pWord=para("Element_Word")
		wPercolor=para("Elemen_Word_Color")
		wPercolorname=para("Element_Word_Color_Name") 
		wPersize=para("Element_Word_Size")
		wPertype=para("Element_Word_Type") 
		pw = para("Element_pic_width")
		ph = para("Element_pic_height")
	  para.close
	end if
	 
	if perLinkStatus=nil then
		perLinkStatus="noLink"
	end if
	if peralign=nil then
		peralign="center"
	end if
	if perTextImgAlign=nil or perTextImgAlign="" then
		perTextImgAlign="center"
	end if
	if wPersize=nil then
		wPersize="2"
	end if
%>
<body topmargin=0 leftmargin=0 rightmargin=0 bottommargin=0>
<table bgcolor="#c0c0c0" border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="5" class="page_title">עדכון טקסט מבצע</td>
  </tr> 
</table>
<table border="0" width="100%"  cellspacing="1" cellpadding="0">
	<tr>
		<td align="center"colspan=2 width=100%>&nbsp;</td>
			</tr>
</table>
<table width="710" cellspacing="1" cellpadding="0" align=center border="0" bgcolor="#ffffff" >

<tr>
    <td align="center" valign="middle" class="small_table_header" height=20 bgcolor="#3E596E" nowrap>עדכון תמונה</td>
</tr>
<tr>
	<td align="right" height="15" bgcolor="#DDDDDD">&nbsp;</td>
</tr>
<%if elemId<>nil then%>
<tr>
	<td align="<%=peralign%>" bgcolor="#DDDDDD">
		<%if perpicture>0  then %>
			<img src="../../GetImage.asp?DB=Template_Element&FIELD=Element_Picture&ID=<%=elemId%>" border="0">
		<%else%>
			&nbsp;
		<%end if%>
	</td>
</tr>
<%end if%>
<form name="form1" ACTION="Aimgadd.asp?C=1&amp;F=Element_Picture&amp;elemId=<%=elemId%>&amp;place=<%=place%>&amp;pageId=<%=pageId%>" METHOD="post" ENCTYPE="multipart/form-data">
<tr>
	<td align="center" bgcolor="#DDDDDD">
		<input TYPE="FILE" NAME="UploadFile1" SIZE="45">
	</td>
</tr>
<tr>
	<td align="center" bgcolor="#DDDDDD">
	  <font face="Arial (Hebrew)">
		<input type="submit" class="but" style="width:100" value="העלאת תמונה" onClick="return ifFieldEmpty()">
	  </font>
	</td>
</tr>
<tr>
	<td height="5" width="100%" bgcolor="#DDDDDD"><table><tr><td></td></tr></table></td>
</tr>
</form>



<%if perpicture>0 then%>
<form name="f1" action="event.asp?pageId=<%=pageId%>&elemId=<%=elemId%>" method="POST">
<tr>
	<td valign="top" bgcolor="white" width="100%">
	<table border="0" cellspacing="0" cellpadding="0" width="100%">
	<tr>
	<td width="50%" bgcolor="#DDDDDD" valign="top">
    <table bgcolor="#DDDDDD" border="0" cellspacing="0" cellpadding="1" width="100%">
    <tr><td colspan="3" align="center" class="small_table_header" bgcolor="#3E596E">קישור</td></tr>	
  <tr>
  	<td colspan="3" bgcolor="#DDDDDD">
  	<table width="100%">
  	<tr>
  	<td  bgcolor="#DDDDDD" align="right" class="form">
		<input type="radio" name="linkTarget" value="_blank" <%if linkTarget="_blank" then%>checked<%end if%> >
		<strong> קישור חיצוני</strong>	
	</td>	
  	<td  bgcolor="#DDDDDD" align="right" class="form">
		<input type="radio" name="linkTarget" value="_self" <%if linkTarget="_self" then%>checked<%end if%> >
	    <strong> קישור פנימי</strong>
	</td>	
  	<td  bgcolor="#DDDDDD" align="right" class="form">
		<input type="radio" name="isLink" value="noLink" <%if perLinkStatus="noLink" then%>checked<%end if%> >
		<strong>ללא קישור</strong>
	</td>
	</tr>
	</table>
	</td>
  </tr>
  <tr>
		<td bgcolor="#DDDDDD" align="right" nowrap>&nbsp;<input type="text" name="allLink" size="30" <%if perLinkStatus="global" then%>value="<%=perLink%>"<%end if%>></td>
		<td bgcolor="#DDDDDD" align="right" nowrap class="form">:קישור ל</td>
		<td bgcolor="#DDDDDD" align="right" nowrap><input type="radio" name="isLink" value="global" <%if perLinkStatus="global" then%>checked<%end if%>></td>
	</tr>
	<tr><td colspan="3" height="5" bgcolor="#DDDDDD"><table><tr><td><td></tr></table></td></tr>
</table>	
	</td>
	<td width="1" nowrap bgcolor="#FFFFFF"></td>
	<td width="50%" bgcolor="#DDDDDD" valign="top">
	<table bgcolor="#DDDDDD" border="0" cellspacing="0" cellpadding="1" width="100%">
	<tr>
		<td colspan="3" align="center" class="small_table_header" bgcolor="#3E596E">יישור</td>
	</tr>	
	<tr>
	    <td align="center" bgcolor="#DDDDDD" width="50%">			
			<select class="app" name="falign" size="3" style="width:80" dir="rtl">
				<option value="center" <%if peralign="center" then%>selected<%end if%>>ממורכז</option>
				<option value="left" <%if peralign="left" then%>selected<%end if%>>לשמאל</option>
				<option value="right" <%if peralign="right" then%>selected<%end if%>>לימין</option>
			</select>			
		  </td>
		</tr>
		<tr><td height="5" align="center" bgcolor="#DDDDDD"><table><tr><td><td></tr></table></td></tr>
		</table>
		</td>
	</tr>
	<tr>	
	<td width="50%" bgcolor="#DDDDDD" valign="top">
	<table bgcolor="#DDDDDD" border="0" cellspacing="0" cellpadding="1" width="100%">
	<tr>
		<td colspan="3" align="center" class="small_table_header" bgcolor="#3E596E">צבע רקע</td>
	</tr>	
	<tr>
		<td bgcolor="#DDDDDD" align="right" nowrap>
		<input type="text" class="texts" name="bgcolnm" style="width:100" value="<%=perbgcolorname%>">
			<input type="hidden" name="bgcolor" value="<%=perbgcolor%>">
			<input type="hidden" name="bgcolname" value="<%=perbgcolorname%>">
		</td>
		<td align="left" bgcolor="#DDDDDD" nowrap>&nbsp;<A  class="linkFaq" HREF="" ONCLICK="return mywindow2();">צבע רקע</a></td>
	</tr>
	<tr><td height="5" align="center" bgcolor="#DDDDDD"><table><tr><td><td></tr></table></td></tr>
	</table>
	</td>
	<td width="1" nowrap bgcolor="#FFFFFF"></td>
	<td width=50% bgcolor="#DDDDDD">
	<table  valign="top"  border="0" cellspacing="0" cellpadding="1" width="100%">
	<tr>
	   <td  valign="top" colspan="3" align="center" class="small_table_header" bgcolor="#3E596E">
	   גודל תמונה
	   </td>
	</tr>
	<tr>
	    <td  valign="top" colspan="3" height="5" class="td_admin_4_normal">
	    </td>
	</tr>
	<tr>
	   <td  valign="top" align="center" class="td_admin_4" width="100%">
	      <table  valign="top"  border="0" cellspacing="0" cellpadding="0" width="100%">
	          <tr>
	          <td width="100%" align="center">
	          <table  valign="top"  border="0" cellspacing="5" cellpadding="0" width="40%">
	          <tr>
	             <td width="10%"><input type="text" size="4" maxlength="4" name="width_pic" value="<%=pw%>" onKeyPress="GetNumbers()">
	             </td>
	             <td width="*%" align="left" class="form">px רוחב:</td>
	          </tr>
	          <tr>
	             <td width="10%"><input type="text" size="4" maxlength="4" name="height_pic" value="<%=ph%>" onKeyPress="return GetNumbers()">
	             </td>
	             <td width="*%" align="left" class="form">px גובה</td>
	          </tr>
	       </table>
	     </td>
	   </tr>
	  </table>
	</td>
	</tr>
	</table>
	</td>
	</tr>
	</table>
	</td>
</tr>
<!--text-->
<tbody name="text_near_image" id="text_near_image">  	
<tr>
<td>
<table border="0" cellspacing="0" cellpadding="1" width="100%">
<tr>
	<td align="center" width="100%" class="small_table_header" bgcolor="#3E596E">
	טקסט על יד התמונה
	</td>
</tr>
<tr><td bgcolor="#DDDDDD" height=10 nowrap></td></tr>	
<tr bgcolor="#DDDDDD">
	<td valign="top" align="center" WIDTH="100%" bgcolor="#DDDDDD">
	<table cellpadding="0" cellspacing="0" border="0" width="100%">
   <tr>
	<td valign="top" align="center" WIDTH="60%" bgcolor="#DDDDDD">
		<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#DDDDDD">
		  <tr>
			<td valign="top" align="left" width="80%" nowrap>
<%pr=left(pertext,1)
if pr="<" then%>
				<textarea align="left" name="title1" dir="ltr" rows="7" style="width:99%"><%=pertext%></textarea>
<%else%>
				<textarea align="left" name="title1" dir="rtl" rows="7" style="width:99%"><%=pertext%></textarea>
<%end if%>
			</td>
<script>
var titleText=document.all("title1").value;
</script>			
<%'if pr="<" then%>
			<td valign="top" align="left" id='clearHTML_td' width="20%" nowrap bgcolor="#DDDDDD">
			 <table cellpadding="0" cellspacing="0" border="0" width="100%">
				<tr><td nowrap align=center>
				<font size=2 face="Arial (Hebrew)" color="#000000"><b>&nbsp;HTML טקסט</b></font>
				</td></tr>
				<tr><td  align=center>
				<input type=radio name=clearHTML id=clearHTML onclick='clearHTMLfunc()' <%if pr="<" then%> checked <%end if%>>
				</td></tr>
				<tr><td nowrap align=center>
				<font size=2 face="Arial (Hebrew)" color="#000000"><b>&nbsp;טקסט חופשי</b></font>
				</td></tr>
				<tr><td  align=center>
				<input type=radio name=clearHTML id=clearHTML onclick='clearHTMLfunc()' <%if pr<>"<" then%> checked <%end if%>>
				</td></tr>
				</table></td>
<%'else%>
<!--			<td valign="top" align="left" nowrap>
				<font face="Arial (Hebrew)" color="#000000">&nbsp;:טסקט</font>
			<td>-->
<%'end if%>
		  </tr>
		</table>
	</td>

</tr>
<tr>
	<td width="100%" valign="top" bgcolor="#DDDDDD">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" bgcolor="#DDDDDD">
	  <tr>
	  <td width="50%" valign="top">
		<table bgcolor="#DDDDDD" border="0" cellspacing="0" cellpadding="0" width="100%">
			<tr><td colspan="2" align="center" bgcolor="#3E596E" height="20" class="small_table_header">מילה  בתוך  הטקסט</td></tr>
			<tr><td height=5 nowrap></td></tr>
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
				<td width="60%" align="center" valign="middle" bgcolor="#DDDDDD" nowrap class="form"><input type="text" class="texts" style="width:120" name="WWord" value="<%=vFix(pWord)%>">&nbsp;:מילה</td>
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
	<td width="25%" valign="top">		
		<table border="0" width=100% cellpadding="0" cellspacing="0" bgcolor="#DDDDDD">
				<tr><td colspan="2" align="center" bgcolor="#3E596E" height="20" class="small_table_header">יישור טקסט</td></tr>
			    <tr><td height=5 nowrap></td></tr>
				<tr>
				<td align="center" width="100%">				
			    <select name="fTxtAlign" size="3" class="app" style="width:90px" dir="rtl">
				<option value="center" <%if perTextImgAlign="center" then%>selected<%end if%>>ממורכז</option>
				<option value="left" <%if perTextImgAlign="left" then%>selected<%end if%>>לשמאל</option>
				<option value="right" <%if perTextImgAlign="right" then%>selected<%end if%>>לימין</option>
			    </select>			
				</td>
				</tr>
		</table>
    </td>
    <td width="1" nowrap bgcolor="#FFFFFF"></td>
	<td width="25%"  valign="top" bgcolor="#DDDDDD">
		<table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="#DDDDDD">
		    <tr><td colspan="2" align="center" bgcolor="#3E596E" height="20" class="small_table_header">גופן</td></tr>
			<tr><td height=5 nowrap></td></tr>
			<tr>
				<td align="right" width=55% nowrap>
				<input type="text" class="texts" name="fcolnm" style="width:95%" value="<%=percolorname%>">
				<input type="hidden" name="fcolor" value="<%=percolor%>">
				<input type="hidden" name="fcolname" value="<%=percolorname%>">
				</td>
				<td align="left" width=45%>&nbsp;<a HREF="" class="linkFaq" ONCLICK="return mywindow();">צבע טקסט</a></td>
			</tr>
			<tr>
				<td align="right" nowrap>
				   <select class="app" style="width:95%" name="fsize" size="1"> 
					<%i=1
					  do while i<=7 %>
    <option value="<%=i%>" <%if persize=i then%>selected<%end if%>><%=i%></option>
					<%i=i+1
					  Loop%>
				   </select>
				</td>
				<td align="left"><font face="Arial (Hebrew)" size=2>&nbsp;:גודל</font></td>
			</tr>
			<tr>
				<td align="right" nowrap>
   					<font face="Arial (Hebrew)">
   					<div dir="rtl">
					<select class="app" style="width:95%" name="fstyle" size="1"> 
						<option value="Regular">רגיל</option>
						<option value="I" <%if pertype="I" then%>selected<%end if%>>נטוי</option>
						<option value="STRONG" <%if pertype="STRONG" then%>selected<%end if%>>מודגש</option>
						<option value="U" <%if pertype="U" then%>selected<%end if%>>קו תחתון</option>
						<option value="SUB" <%if pertype="SUB" then%>selected<%end if%>>כתב עילי</option>
						<option value="SUP" <%if pertype="SUP" then%>selected<%end if%>>כתב תחתי</option>
						<option value="STRIKE" <%if pertype="STRIKE" then%>selected<%end if%>>קו חוצה</option>
					</select></div>
					</font>
				</td>
				<td align="left"><font face="Arial (Hebrew)" size=2>&nbsp;:סגנון</font></td>
			</tr>
			<!--tr>
				<td align="right" nowrap>
					<%i=5%>					
					<select class="app" style="width:95%" name="hebrSize" size="1">
						<option value="null">חופשי</option>
						<%do while i<=100%>
						<option value="<%=i%>" <%if perHebrSize=i then%>selected<%end if%>><%=i%></option>
						<%i=i+1
					      loop%>
					</select> 					
				</td>
				<td align="left">&nbsp;<font face="Arial (Hebrew)" size=2>:הרושה ךרוא</font></td>
			</tr-->
		</table>	
	   </td>			
		</tr>
		</table>	
		</td>
	  </tr>	  
	  </table>
	</td>
</tr>
</table></td></tr>
</tbody>
<tr><!--HTML-->
	<td align="right">	
   <table border="0" width="100%" align="center" cellspacing="0" cellpadding="1">
  <tr>
    <td align="center" valign="middle" class="small_table_header" bgcolor="#3E596E" nowrap>HTML</td>
  </tr>
<!--<form name="f2" ACTION="event.asp?innerparent=<%=innerparent%>&subcat=<%=subcat%>&isright=<%=isRightMenu%>" METHOD="post">-->
<tr>
	<td align="right" bgcolor="#DDDDDD">
	<table  border=0 height="20" border=0 bgcolor="#DDDDDD">
	<tr>
	<td nowrap colspan=11 align=right class="form">
	<b> בטבלה HTML לבנות טקסט</b>
	<input type="checkbox" name="numcolumn" LANGUAGE=javascript onclick="return numcolumn_onclick('<%=pr%>')">
	</td>  
	</tr>
	<tr id=textpict style="display:none">
	<td>
		<INPUT type="button" value="בנה טבלה" class="but" LANGUAGE=javascript onclick="return button1_onclick()">
	</td>
	<td width=10>
		&nbsp;
	</td>
	<td nowrap class="form">
		<SELECT class="app" style="width:50" id=numborder name=numborder>
		<%for i=0 to 6%>		<OPTION value=<%=i%>><%=i%></OPTION>
		<%next%>		</SELECT>
		<b>גבול</b>
	</td>
	<td width=10>
		&nbsp;
	</td>
	<td nowrap class="form">
		<SELECT class="app" style="width:70" id=numwidth name=numwidth>
		<option value=320>320px</OPTION>
		<option value=300>300px</OPTION>
		<option value=250>250px</OPTION>
		<option value=250>200px</OPTION>
		<%'for i=10 to 100 step 5%>		<!--OPTION value=<%=i%> <%if i=100 then%> selected<%end if%>><%=i%> %</OPTION-->
		<%'next%>		</SELECT>
		<b>רוחב טבלה</b>
	</td>
	<td width=10>
		&nbsp;
	</td>
	<td nowrap class="form">
		<SELECT class="app" style="width:50" id=columnnumb name=columnnumb>
		<%for i=1 to 4%>		<OPTION value=<%=i%>><%=i%></OPTION>
		<%next%>		</SELECT>
		<b>כמות טורים</b>
	</td>
	<td width=10>
		&nbsp;
	</td>
	<td nowrap class="form">
		<SELECT class="app" style="width:50" id=numrow name=numrow>
		<%for i=1 to 30%>		<OPTION value=<%=i%>><%=i%></OPTION>
		<%next%>		</SELECT>
		<b>כמות שורות</b>
	</td>
	<td width=10>
		&nbsp;
	</td>
	<td nowrap align=left class="form">
		<SELECT class="app" style="width:70" id=tbl_align name=tbl_align>
			<OPTION value=right>לימין</OPTION>			<OPTION value=center>ממורכז</OPTION>			<OPTION value=left>לשמאל</OPTION>
		</SELECT>
		<b>ישר טבלה</b>
	</td>
	</tr>
	</table>
	</td>
  </tr>

  <tr bgcolor="#DDDDDD">
	 <td align=right><div id=mytable name=mytable></div>
	 <!--font face="Arial (Hebrew)"><textarea name="htmltext2" rows="15" cols="70"><%'=pertext2%></textarea></font-->
	 </td>
</tr>

<!--//////////////////////////////////////////////-->
	    <input type="hidden" name="elemId" value="<%=elemId%>">
	    <input type="hidden" name="pageId" value="<%=pageId%>">
	    <input type="hidden" name="prodId" value="<%=prodId%>">
	    <input type="hidden" name="homePage" value="<%if perHomePage=True then%>&quot;1&quot;<%else%>&quot;0&quot;<%end if%>">
	    <input type="hidden" name="place" value="<%=place%>">
	    <input type="hidden" name="pictureHTML"  <%if pr="<" then%>value="1"<%else%> value=""<%end if%>>
	    <input type="hidden" name="htmltext2" value="">

</table>	
	</td>
</tr>
<!--input type="hidden" name="elemId" VALUE="<%=elemId%>"--> 
<!--input type="hidden" name="pageId" VALUE="<%=pageId%>"--> 
<!--input type="hidden" name="homePage" VALUE="<%if perHomePage=True then%>&quot;1&quot;<%else%>&quot;0&quot;<%end if%>"--> 
<input type="hidden" name="editPicture" value="1">
<input type="hidden" name="stam" value="<%=perpicture%>">
</form>

<tr><td>
<table border="0" width="100%" align="center" cellspacing="1" cellpadding="0" bgcolor="#DDDDDD">
<tr>
	<td width="100%" nowrap bgcolor="#DDDDDD">
	<table width=100% border="0" cellpadding="1" cellspacing="0" bgcolor="#DDDDDD">
	  <tr height=30 valign="middle">
		<td bgcolor="#DDDDDD" width=45% nowrap align="right">
	    <INPUT class="but" style="width:90"  type="button" onclick="document.location.href='event.asp?innerparent=<%=innerparent%>&pageId=<%=pageId%>&amp;prodId=<%=prodId%>&isright=<%=isRightMenu%>'" value="ביטול" id=button3 name=button3>
		</td>
		<td bgcolor="#DDDDDD" width=5% nowrap align="left"></tD>
		<td bgcolor="#DDDDDD" width=45% nowrap align="left">	    
	    <INPUT class="but" style="width:90" type="button" onclick="return FormSubmitFun();" value="אישור" id=button2 name=button2>
		</td>		
	  </tr>
	</table>
	</td>
</tr>
<%end if%>
</table><!--*-->
</div>
</body>
<%set con=Nothing%>
</html>
