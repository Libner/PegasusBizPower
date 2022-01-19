<%
'on error resume next
Function CheckType(i,mass,eType,elemId,eNumOnPage,eOrderAlign,eText,eLinkStatus,eLinkTarget,eLink,eBullet,eWWord,eWColor,eWColorName,eWSize,eWType,eFileName,ePictureSize,eFontSize,eFontType,eFontColor,eBGColor,eHebrSize,eAlign,eHomePage,eConnection,eVis)
	Select Case eType
		case "1"
			textFunction i,mass,elemId,eNumOnPage,eOrderAlign,eText,eLinkStatus,eLinkTarget,eLink,eBullet,eWWord,eWColor,eWColorName,eWSize,eWType,eFontSize,eFontType,eFontColor,eBGColor,eHebrSize,eAlign,eHomePage,eConnection,eVis
		case "2"
			imageFunction i,mass,elemId,eNumOnPage,eOrderAlign,eText,eLinkStatus,eLinkTarget,eLink,ePictureSize,eFontSize,eFontType,eFontColor,eHebrSize,eAlign,eTxtImgAlign,eHomePage,eConnection,eVis
		case "3"
			fileFunction i,mass,elemId,eNumOnPage,eOrderAlign,eText,eFileName,eFontSize,eFontType,eFontColor,eHebrSize,eAlign,eHomePage,eConnection,eVis
		case "4"
			mailFunction i,mass,elemId,eNumOnPage,eOrderAlign,eText,eFileName,eFontSize,eFontType,eFontColor,eHebrSize,eAlign,eHomePage,eConnection,eVis
		case "5"
			HTMLFunction i,mass,elemId,eNumOnPage,eOrderAlign,eText,eFileName,eFontSize,eFontType,eFontColor,eHebrSize,eAlign,eHomePage,eConnection,eVis
	End Select
End Function	

Function textFunction(i,mass,elemId,eNumOnPage,eOrderAlign,eText,eLinkStatus,eLinkTarget,eLink,eBullet,eWWord,eWColor,eWColorName,eWSize,eWType,eFontSize,eFontType,eFontColor,eBGColor,eHebrSize,eAlign,eHomePage,eConnection,eVis)
        if isNull(eText) then
           eText=" "
        end if   
		if eHebrSize<>nil then	
			eTextH=hebrew(eText,eHebrSize)
		else		
		  koeff=160\eFontSize	
		   	eTextH=hebrew(eText,koeff)
		end if

if eLinkStatus="global" then
       'eTextH="<a href=http://" & eLink & " target=_blank><font color=" & eFontColor & ">" & eTextH & "</font></a>"		
       eTextH="<a href=" & eLink & " target=" & eLinkTarget & "><font color=" & eFontColor & ">" & eTextH & "</font></a>"		
end if
if eLinkStatus="words" then
  set lin=con_net.execute("SELECT * from Link_Texts where Element_Id="&elemId&" ")
  s="SELECT * from Link_Texts where Element_Id="&elemId&" "
  'response.write(s)
  if not lin.eof then
		key_txt=lin("Key_Text")
			if eHebrSize<>nil then	
				key_txt=hebrew(key_txt,eHebrSize)
			else		
			  koeff=160\eFontSize	
			  key_txt=hebrew(key_txt,koeff)
			end if
		ln_txt=lin("Link_Text")
	   'rep_txt="<a href=http://" & ln_txt & "><font color=#0000ff>" & key_txt & "</font></a>"
       rep_txt="<a href=" & ln_txt & " target=" & eLinkTarget & "><font color=#0000ff>" & key_txt & "</font></a>"
       eTextH=replace(eTextH,key_txt,rep_txt)		
  end if
  lin.close
end if

			if eWWord<>nil then
	if len(Trim(eWWord))>0 then
		eWWordH=reverse(eWWord)
		tmpStr="<font style=font-size:" & 7+eWSize*2 & "pt; color=" & eWColor & "><" & eWType & ">" & eWWordH & "</" & eWType & "></font>"
		eTextH=replace(eTextH,eWWordH,tmpStr)
	end if
			end if	
%>
<tr>
	<td width="3%" align="center" valign="bottom" bgcolor="#dddddd"><a href="editText.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;pageId=<%=pageId%>"><img src="images/update.gif" alt="לעדכן את הנתוני הטקסט" border="0" width="56"></a></td>
	<%'page main%>
	<!--<td rowspan="2" align="center" valign="middle" width="1%" bgcolor="#dddddd" nowrap>
	    <%if eHomePage=True then%>
	     <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;lamp=off"><img src="images/Main_lamp_on.gif" alt="עכשיו הטקסט מופיע בדף ראשי של הקטגוריה" border="0"></a>
		<%else%>
		 <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;lamp=on"><img src="images/Main_lamp_off.gif" alt="עכשיו הטקסט לא מופיע בדף ראשי של הקטגוריה" border="0"></a>
		<%end if%>
	</td>
	<td rowspan="2" align="center" valign="middle" width="1%" bgcolor="#dddddd" nowrap>
	    <%if eConnection=True then%>
	     <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;conn=off"><img src="images/Arrow_on.gif" alt="עכשיו הטקסט הוא קישור לדף פנימי" border="0"></a>
		<%else%>
	     <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;conn=on"><img src="images/Arrow_off.gif" alt="עכשיו הטקסט הוא קישור לדף פנימי" border="0"></a>
		<%end if%>
	</td>
	<%'end page main%>
	<td rowspan="2" align="center" valign="middle" width="1%" bgcolor="#dddddd" nowrap>
	    <%if eVis=True then%>
	     <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;vis=off"><img src="images/My_lamp_on.gif" alt="עכשיו הטקסט מופיע בדף פנימי" border="0" width="34"></a>
		<%else%>
	     <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;vis=on"><img src="images/My_lamp_off.gif" alt="עכשיו הטקסט לא מופיע בדף פנימי" border="0" width="34"></a>
		<%end if%>
	</td>-->
    <td rowspan="2" align="<%=eAlign%>" valign="middle" bgcolor="#feffe7">
     <table border="0" cellpadding="0" cellspacing="0" width=100%>
     <tr>
      <td align="<%=eAlign%>" valign="middle" bgcolor="#feffe7" nowrap>
      <font color="<%=eFontColor%>" style="font-size:<%=7+eFontSize*2%>pt;">
	   <%if eFontType<>nil then%><<%=eFontType%>><%end if%>
		<%=eTextH%>
       <%if eFontType<>nil then%></<%=eFontType%>><%end if%>
      </font>
      </td>
      <td width="5" valign="top" bgcolor="#feffe7">
      <%if not isNull(eBullet) and eBullet<>"0" then%>
		<img src="../../images/bul<%=eBullet%>.gif" border="0">
      <%else%>
		&nbsp;
      <%end if%>
      </td>
      </tr>
      </table>
    </td>
	<td width="1%" align="center" valign="bottom" bgcolor="#dddddd">
	<%if i>1 then%>
		<a href="event.asp?isright=<%=isRightMenu%>&pageId=<%=pageId%>&amp;up=1&amp;place=<%=eNumOnPage%>"><img src="images/up.gif" alt="להעביר את הטקסט למעלה" border="0" width="11"></a>
	<%else%>
	  &nbsp;
	<%end if%>
	</td>
</tr>
<tr>
	<td width="3%" align="center" valign="top" bgcolor="#dddddd">
<a href="event.asp?isright=<%=isRightMenu%>&DEL=1&amp;elemId=<%=elemId%>&amp;pageId=<%=pageId%>&amp;place=<%=eNumOnPage%>" ONCLICK="return CheckDel()"><img src="images/delete.gif" alt="למחוק את הטקסט" border="0" width="56"></a>
	</td>
	<td width="1%" align="center" valign="top" bgcolor="#dddddd">
	<%if i<mass then%>
		<a href="event.asp?isright=<%=isRightMenu%>&pageId=<%=pageId%>&amp;down=1&amp;place=<%=eNumOnPage%>"><img src="images/down.gif" alt="להעביר את הטקסט למטה" border="0" width="11"></a>
	<%else%>
	  &nbsp;
	<%end if%>
	</td>
</tr>
<%
End Function

Function imageFunction(i,mass,elemId,eNumOnPage,eOrderAlign,eText,eLinkStatus,eLinkTarget,eLink,ePictureSize,eFontSize,eFontType,eFontColor,eHebrSize,eAlign,eTxtImgAlign,eHomePage,eConnection,eVis)
	if not isNull(eText) then
		if eHebrSize<>nil then	
			eTextH=hebrew(eText,eHebrSize)
		else		
		  koeff=120\eFontSize	
		   	eTextH=hebrew(eText,koeff)
		end if
	end if
if eTextH=nil or eTextH="" then
	eTextH=" "
end if
%>
<tr>
	<td width="3%" bgcolor="#dddddd" valign="bottom" align="center">
	   <a href="AddPicture.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;pageId=<%=pageId%>"><img src="images/update.gif" alt="לעדכן את הנתוני התמונה" border="0" width="56"></a>
	</td>
	<!--<td rowspan="2" align="center" valign="middle" width="1%" bgcolor="#dddddd" nowrap>
	    <%if eHomePage=True then%>
	     <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;lamp=off"><img src="images/Main_lamp_on.gif" alt="עכשיו התמונה מופיעה בדף ראשי של הקטגוריה" border="0"></a>
		<%else%>
		 <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;lamp=on"><img src="images/Main_lamp_off.gif" alt="עכשיו התמונה לא מופיעה בדף ראשי של הקטגוריה" border="0"></a>
		<%end if%>
	</td>
	<td rowspan="2" align="center" valign="middle" width="1%" bgcolor="#dddddd" nowrap>
	    <%if eConnection=True then%>
	     <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;conn=off"><img src="images/Arrow_on.gif" alt="עכשיו התמונה היא קישור לדף פנימי" border="0"></a>
		<%else%>
	     <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;conn=on"><img src="images/Arrow_off.gif" alt="עכשיו התמונה היא לא קישור לדף פנימי" border="0"></a>
		<%end if%>
	</td>
	<td rowspan="2" align="center" valign="middle" width="1%" bgcolor="#dddddd" nowrap>
	    <%if eVis=True then%>
	     <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;vis=off"><img src="images/My_lamp_on.gif" alt="עכשיו התמונה מופיע בדף פנימי" border="0" width="34"></a>
		<%else%>
	     <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;vis=on"><img src="images/My_lamp_off.gif" alt="עכשיו התמונה לא מופיע בדף פנימי" border="0" width="34"></a>
		<%end if%>
	</td>-->
	<td rowspan="2" align="<%=eAlign%>" valign="top" bgcolor="#feffe7">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr>
	<%if eAlign="right" and eTextH<>nil then%>
<td bgcolor="#feffe7" valign="top" align="<%=eTxtImgAlign%>">
      <font color="<%=eFontColor%>" style="font-size:<%=7+eFontSize*2%>pt;">
	   <%if eFontType<>nil then%><<%=eFontType%>><%end if%>
		<%=eTextH%>
       <%if eFontType<>nil then%></<%=eFontType%>><%end if%>
      </font>
</td>
	<%end if%>
<td bgcolor="#feffe7" valign="top" align="<%=eAlign%>" width="10%">
		<%if ePictureSize>0 then %>
<%if eLinkStatus="global" then%><a href="<%=eLink%>" target=<%=eLinkTarget%>><%end if%><img src="../../GetImage.asp?DB=Element&amp;FIELD=Element_Picture&amp;ID=<%=elemId%>" border="0" vspace="5" hspace="5"><%if eLinkStatus="global" then%></a><%end if%>
		<%else%>&nbsp;<%end if%>
</td>
	<%if eAlign="left" and eTextH<>nil then%>
<td bgcolor="#feffe7" valign="top" align="<%=eTxtImgAlign%>">
      <font color="<%=eFontColor%>" style="font-size:<%=7+eFontSize*2%>pt;">
	   <%if eFontType<>nil then%><<%=eFontType%>><%end if%>
		<%=eTextH%>
       <%if eFontType<>nil then%></<%=eFontType%>><%end if%>
      </font>
</td>
	<%end if%>
	</tr>
<%if eAlign="center" and eTextH<>nil then%>	
	<tr>
		<td bgcolor="#feffe7" align="<%=eTxtImgAlign%>">
      <font color="<%=eFontColor%>" style="font-size:<%=7+eFontSize*2%>pt;">
	   <%if eFontType<>nil then%><<%=eFontType%>><%end if%>
		<%=eTextH%>
       <%if eFontType<>nil then%></<%=eFontType%>><%end if%>
      </font>
		</td>
	</tr>
<%end if%>
</table>
    </td>
	<td width="1%" align="center" valign="bottom" bgcolor="#dddddd">
	<%if i>1 then%>
		<a href="event.asp?isright=<%=isRightMenu%>&pageId=<%=pageId%>&amp;up=1&amp;place=<%=eNumOnPage%>"><img src="images/up.gif" alt="להעביר את התמונה למעלה" border="0" width="11"></a>
	<%else%>
	  &nbsp;
	<%end if%>
	</td>
</tr>
<tr>
	<td width="3%" align="center" valign="top" bgcolor="#dddddd">
<a href="event.asp?isright=<%=isRightMenu%>&DEL=1&amp;elemId=<%=elemId%>&amp;pageId=<%=pageId%>&amp;place=<%=eNumOnPage%>" ONCLICK="return CheckDel()"><img src="images/delete.gif" alt="למחוק את התמונה" border="0" width="56"></a>
	</td>
	<td width="1%" align="center" valign="top" bgcolor="#dddddd">
	<%if i<mass then%>
		<a href="event.asp?isright=<%=isRightMenu%>&pageId=<%=pageId%>&amp;down=1&amp;place=<%=eNumOnPage%>"><img src="images/down.gif" alt="להעביר את התמונה למטה" border="0" width="11"></a>
	<%else%>
	  &nbsp;
	<%end if%>
	</td>
</tr>
<%
End Function

Function fileFunction(i,mass,elemId,eNumOnPage,eOrderAlign,eText,eFileName,eFontSize,eFontType,eFontColor,eHebrSize,eAlign,eHomePage,eConnection,eVis)
        if isnull(eText) then
           eText=" "
         end if  
		if eHebrSize<>nil then	
			eTextH=hebrew(eText,eHebrSize)
		else		
		  koeff=260\eFontSize	
		   	eTextH=hebrew(eText,koeff)
		end if
if eFileName<>nil then
extend=LCase(right(eFileName,len(eFileName)-InstrRev(eFileName,".")))
end if
%>
<tr>
	<td width="3%" align="center" valign="bottom" bgcolor="#dddddd"><a href="addFile.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;pageId=<%=pageId%>"><img src="images/update.gif" alt="לעדכן את הנתוני הקובץ" border="0" width="56"></a></td>
	<!--<td rowspan="2" align="center" valign="middle" width="1%" bgcolor="#dddddd" nowrap>
	    <%if eHomePage=True then%>
	     <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;lamp=off"><img src="images/Main_lamp_on.gif" alt="עכשיו הקובץ מופיע בדף ראשי של הקטגוריה" border="0" width="34"></a>
		<%else%>
		 <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;lamp=on"><img src="images/Main_lamp_off.gif" alt="עכשיו הקובץ לא מופיע בדף ראשי של הקטגוריה" border="0" width="34"></a>
		<%end if%>
	</td>
	<td rowspan="2" align="center" valign="middle" width="1%" bgcolor="#dddddd" nowrap>
	     <img src="images/Arrow_off.gif" border="0">
	</td>
	<td rowspan="2" align="center" valign="middle" width="1%" bgcolor="#dddddd" nowrap>
	    <%if eVis=True then%>
	     <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;vis=off"><img src="images/My_lamp_on.gif" alt="עכשיו הקובץ מופיע בדף שלי" border="0" width="34"></a>
		<%else%>
	     <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;vis=on"><img src="images/My_lamp_off.gif" alt="עכשיו הקובץ לא מופיע בדף שלי" border="0" width="34"></a>
		<%end if%>
	</td>-->
    <td rowspan="2" align="<%=eAlign%>" valign="top" bgcolor="#feffe7">
		<table border="0" cellpadding="0" cellspacing="0">
		<tr><td valign="top">
       <%if eFontType<>nil then%><<%=eFontType%>><%end if%>
		<a href="../../download_h/<%=eFileName%>" target="_new"><font color="<%=eFontColor%>" style="font-size:<%=7+eFontSize*2%>pt;"><%if eTextH=nil or eTextH="" then%><%=eFileName%><%else%><%=eTextH%><%end if%></font>
		</a>
       <%if eFontType<>nil then%></<%=eFontType%>><%end if%>
		</td>
		<td width="5" valign="top">
		<a href="../../download_h/<%=eFileName%>" target="_new">
		<%if extend="doc" then%>
			<img src="images/word.gif" border="0">
		<%else%>
			<%if extend="xls" then%>
				<img src="images/excel.gif" border="0">
			<%else%>
				<%if extend="htm" or extend="html" then%>
				<img src="images/html.gif" border="0">
				<%else%>
				<img src="images/disk.gif" border="0">
				<%end if%>
			<%end if%>
		<%end if%>	
		</a>
		</td>
       </tr>
       </table>
    </td>
	<td width="1%" align="center" valign="bottom" bgcolor="#dddddd">
	<%if i>1 then%>
		<a href="event.asp?isright=<%=isRightMenu%>&pageId=<%=pageId%>&amp;up=1&amp;place=<%=eNumOnPage%>"><img src="images/up.gif" alt="להעביר את הקובץ למעלה" border="0" width="11"></a>
	<%else%>
	  &nbsp;
	<%end if%>
	</td>
</tr>
<tr>
	<td width="3%" align="center" valign="top" bgcolor="#dddddd">
<a href="event.asp?isright=<%=isRightMenu%>&DEL=1&amp;elemId=<%=elemId%>&amp;pageId=<%=pageId%>&amp;place=<%=eNumOnPage%>" ONCLICK="return CheckDel()"><img src="images/delete.gif" alt="למחוק את הקובץ" border="0" width="56"></a>
	</td>
	<td width="1%" align="center" valign="top" bgcolor="#dddddd">
	<%if i<mass then%>
		<a href="event.asp?isright=<%=isRightMenu%>&pageId=<%=pageId%>&amp;down=1&amp;place=<%=eNumOnPage%>"><img src="images/down.gif" alt="להעביר את הקובץ למטה" border="0" width="11"></a>
	<%else%>
	  &nbsp;
	<%end if%>
	</td>
</tr>
<%
End Function

Function mailFunction(i,mass,elemId,eNumOnPage,eOrderAlign,eText,eFileName,eFontSize,eFontType,eFontColor,eHebrSize,eAlign,eHomePage,eConnection,eVis)
        if isNull(eText) then
           eText=" "
         end if  
		if eHebrSize<>nil then	
			eTextH=hebrew(eText,eHebrSize)
		else		
		  koeff=220\eFontSize	
		   	eTextH=hebrew(eText,koeff)
		end if
%>
<tr>
	<td width="3%" align="center" valign="bottom" bgcolor="#dddddd"><a href="addMail.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;pageId=<%=pageId%>"><img src="images/update.gif" alt="לעדכן את הנתוני הקובץ" border="0" width="56"></a></td>
	<!--<td rowspan="2" align="center" valign="middle" width="1%" bgcolor="#dddddd" nowrap>
	    <%if eHomePage=True then%>
	     <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;lamp=off"><img src="images/Main_lamp_on.gif" alt="עכשיו הקובץ מופיע בדף ראשי של הקטגוריה" border="0" width="34"></a>
		<%else%>
		 <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;lamp=on"><img src="images/Main_lamp_off.gif" alt="עכשיו הקובץ לא מופיע בדף ראשי של הקטגוריה" border="0" width="34"></a>
		<%end if%>
	</td>
	<td rowspan="2" align="center" valign="middle" width="1%" bgcolor="#dddddd" nowrap>
	     <img src="images/Arrow_off.gif" border="0">
	</td>
	<td rowspan="2" align="center" valign="middle" width="1%" bgcolor="#dddddd" nowrap>
	    <%if eVis=True then%>
	     <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;vis=off"><img src="images/My_lamp_on.gif" alt="עכשיו הקובץ מופיע בדף שלי" border="0" width="34"></a>
		<%else%>
	     <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;vis=on"><img src="images/My_lamp_off.gif" alt="עכשיו הקובץ לא מופיע בדף שלי" border="0" width="34"></a>
		<%end if%>
	</td>-->
    <td rowspan="2" align="<%=eAlign%>" valign="top" bgcolor="#feffe7">
		<table border="0" cellpadding="0" cellspacing="0">
		<tr><td valign="top">
       <%if eFontType<>nil then%><<%=eFontType%>><%end if%>
		<a href="mailto:<%=eFileName%>"><font color="<%=eFontColor%>" style="font-size:<%=7+eFontSize*2%>pt;"><%if eTextH=nil or eTextH="" then%><%=eFileName%><%else%><%=eTextH%><%end if%></font>
		</a>
       <%if eFontType<>nil then%></<%=eFontType%>><%end if%>
		</td>
       </tr>
       </table>
    </td>
	<td width="1%" align="center" valign="bottom" bgcolor="#dddddd">
	<%if i>1 then%>
		<a href="event.asp?isright=<%=isRightMenu%>&pageId=<%=pageId%>&amp;up=1&amp;place=<%=eNumOnPage%>"><img src="images/up.gif" alt="להעביר את הקובץ למעלה" border="0" width="11"></a>
	<%else%>
	  &nbsp;
	<%end if%>
	</td>
</tr>
<tr>
	<td width="3%" align="center" valign="top" bgcolor="#dddddd">
<a href="event.asp?isright=<%=isRightMenu%>&DEL=1&amp;elemId=<%=elemId%>&amp;pageId=<%=pageId%>&amp;place=<%=eNumOnPage%>" ONCLICK="return CheckDel()"><img src="images/delete.gif" alt="למחוק את הקובץ" border="0"></a>
	</td>
	<td width="1%" align="center" valign="top" bgcolor="#dddddd">
	<%if i<mass then%>
		<a href="event.asp?isright=<%=isRightMenu%>&pageId=<%=pageId%>&amp;down=1&amp;place=<%=eNumOnPage%>"><img src="images/down.gif" alt="להעביר את הקובץ למטה" border="0" width="11"></a>
	<%else%>
	  &nbsp;
	<%end if%>
	</td>
</tr>
<%
End Function
%>

<%
Function HTMLFunction(i,mass,elemId,eNumOnPage,eOrderAlign,eText,eFileName,eFontSize,eFontType,eFontColor,eHebrSize,eAlign,eHomePage,eConnection,eVis)
        if isNull(eText) then
           eText=" "
         end if  
%>
<tr>
	<td width="3%" align="center" valign="bottom" bgcolor="#dddddd"><a href="editHTML.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;pageId=<%=pageId%>"><img src="images/update.gif" alt="לעדכן את הנתוני הקובץ" border="0" width="56"></a></td>
	<!--<td rowspan="2" align="center" valign="middle" width="1%" bgcolor="#dddddd" nowrap>
	    <%if eHomePage=True then%>
	     <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;lamp=off"><img src="images/Main_lamp_on.gif" alt="עכשיו הקובץ מופיע בדף ראשי של הקטגוריה" border="0" width="34"></a>
		<%else%>
		 <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;lamp=on"><img src="images/Main_lamp_off.gif" alt="עכשיו הקובץ לא מופיע בדף ראשי של הקטגוריה" border="0" width="34"></a>
		<%end if%>
	</td>
	<td rowspan="2" align="center" valign="middle" width="1%" bgcolor="#dddddd" nowrap>
	     <img src="images/Arrow_off.gif" border="0">
	</td>
	<td rowspan="2" align="center" valign="middle" width="1%" bgcolor="#dddddd" nowrap>
	    <%if eVis=True then%>
	     <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;vis=off"><img src="images/My_lamp_on.gif" alt="עכשיו הקובץ מופיע בדף שלי" border="0" width="34"></a>
		<%else%>
	     <a href="event.asp?isright=<%=isRightMenu%>&elemId=<%=elemId%>&amp;vis=on"><img src="images/My_lamp_off.gif" alt="עכשיו הקובץ לא מופיע בדף שלי" border="0" width="34"></a>
		<%end if%>
	</td>-->
    <td rowspan="2" align="<%=center%>" valign="top" bgcolor="#feffe7">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
		<tr>
		<td valign="top">
       <%=eText%>
		</td>
       </tr>
       </table>
    </td>
	<td width="1%" align="center" valign="bottom" bgcolor="#dddddd">
	<%if i>1 then%>
		<a href="event.asp?isright=<%=isRightMenu%>&pageId=<%=pageId%>&amp;up=1&amp;place=<%=eNumOnPage%>"><img src="images/up.gif" alt="להעביר את הקובץ למעלה" border="0" width="11"></a>
	<%else%>
	  &nbsp;
	<%end if%>
	</td>
</tr>
<tr>
	<td width="3%" align="center" valign="top" bgcolor="#dddddd">
<a href="event.asp?isright=<%=isRightMenu%>&DEL=1&amp;elemId=<%=elemId%>&amp;pageId=<%=pageId%>&amp;place=<%=eNumOnPage%>" ONCLICK="return CheckDel()"><img src="images/delete.gif" alt="למחוק את הקובץ" border="0"></a>
	</td>
	<td width="1%" align="center" valign="top" bgcolor="#dddddd">
	<%if i<mass then%>
		<a href="event.asp?isright=<%=isRightMenu%>&pageId=<%=pageId%>&amp;down=1&amp;place=<%=eNumOnPage%>"><img src="images/down.gif" alt="להעביר את הקובץ למטה" border="0" width="11"></a>
	<%else%>
	  &nbsp;
	<%end if%>
	</td>
</tr>
<%
End Function
%>
