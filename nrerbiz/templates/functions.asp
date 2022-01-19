<%
'on error resume next
Function CheckType(i,mass,eType,elemId,eNumOnPage,eOrderAlign,eText,eLinkStatus,eLinkTarget,eLink,eBullet,eWWord,eWColor,eWColorName,eWSize,eWType,eFileName,ePictureSize,eFontSize,eFontType,eFontColor,eBGColor,eHebrSize,eAlign,eHomePage,eConnection,eVis,ePicWidth,ePicHight)
	Select Case eType
		case "1"
			textFunction i,mass,elemId,eNumOnPage,eOrderAlign,eText,eLinkStatus,eLinkTarget,eLink,eBullet,eWWord,eWColor,eWColorName,eWSize,eWType,eFontSize,eFontType,eFontColor,eBGColor,eHebrSize,eAlign,eHomePage,eConnection,eVis
		case "2"
			imageFunction i,mass,elemId,eNumOnPage,eOrderAlign,eText,eLinkStatus,eLinkTarget,eLink,eBullet,eWWord,ePictureSize,eFontSize,eFontType,eFontColor,eBGColor,eHebrSize,eAlign,eTxtImgAlign,eHomePage,eConnection,eVis,ePicWidth,ePicHight		
		case "5"
			HTMLFunction i,mass,elemId,eNumOnPage,eOrderAlign,eText,eFileName,eFontSize,eFontType,eFontColor,eBGColor,eHebrSize,eAlign,eHomePage,eConnection,eVis
		
	End Select
End Function	

Function textFunction(i,mass,elemId,eNumOnPage,eOrderAlign,eText,eLinkStatus,eLinkTarget,eLink,eBullet,eWWord,eWColor,eWColorName,eWSize,eWType,eFontSize,eFontType,eFontColor,eBGColor,eHebrSize,eAlign,eHomePage,eConnection,eVis)
    if eSpace = nil then
      eSpace = "5"
    end if
    
    if isNull(eText) then
       eText=" "
    end if   
	
	eTextH = eText
	If Len(trim(eBGColor)) = 0 Then
		eBGColor = "#FFFFFF"
	End If	

	if eLinkStatus="global" then
	   If Len(trim(eLink)) > 0 Then
       if inStr(eLink,"http://") < 1 Then
		   eLink =  "http://" & eLink
	   end if	   
       'eTextH="<a href=http://" & eLink & " target=_blank><font color=" & eFontColor & ">" & eTextH & "</font></a>"		
       eTextH="<a href=" & eLink & " target=" & eLinkTarget & "><font color=" & eFontColor & "  style=font-size:" & 6+eFontSize*2 & "pt;>" & eTextH & "</font></a>"		
       End If
	end if
	if eLinkStatus="words" then
	If Len(trim(eLinkTarget)) = 0 Or isNull(eLinkTarget) Then
		eLinkTarget = "'_blank'"
	End If
	
	set lin=con.getRecordSet("SELECT * from Link_Texts where Element_Id="&elemId&" ")
	s="SELECT * from Link_Texts where Element_Id="&elemId&" "
	'response.write(s)
	if not lin.eof then	
		key_txt=lin("Key_Text")
			
		ln_txt=lin("Link_Text")
		If inStr(ln_txt,"http://") = 0 Then
			ln_txt = "http://" & ln_txt
		End If
			
	   'rep_txt="<a href=http://" & ln_txt & "><font color=#0000ff>" & key_txt & "</font></a>"
       rep_txt="<a href=" & ln_txt & " target=" & eLinkTarget & "><font color=#0000ff>" & key_txt & "</font></a>"
       eTextH=replace(eTextH,key_txt,rep_txt)		
  lin.close
  end if
end if


	if len(Trim(eWWord))>0 then
		'eWWordH=con.reverse(eWWord)
		eWWordH=eWWord
		tmpStr="<font style=font-size:" & 6+eWSize*2 & "pt; color=" & eWColor & "><" & eWType & ">" & eWWordH & "</" & eWType & "></font>"
		eTextH=replace(eTextH,eWWordH,tmpStr)
	end if

%>
<tr>
	<td width="3%" align="center" valign="bottom" bgcolor="#dddddd">
	<a href="editText.asp?elemId=<%=elemId%>&pageId=<%=pageId%>&prodId=<%=prodId%>" class="but_edit">עדכון</a></td>
    <td rowspan="2" align="<%=eAlign%>" valign="middle" bgcolor="<%=eBGColor%>">
     <table border="0" cellpadding="0" cellspacing="0" align="<%=eAlign%>">
     <tr>
      <%if (eSpace<>nil) then%>  
	   <td width="<%=eSpace%>" nowrap></td>
	  <%end if%>    
      <td dir="rtl" align="<%=eAlign%>" valign="middle" bgcolor="<%=eBGColor%>">
      <font color="<%=eFontColor%>" style="font-size:<%=6+eFontSize*2%>pt;">
	   <%if eFontType<>nil then%><%=chr(60)%><%=eFontType%><%=chr(62)%><%end if%>
		<%=breaks(eTextH)%>
       <%if eFontType<>nil then%><%=chr(60)%><%=chr(47)%><%=eFontType%><%=chr(62)%><%end if%>
      </font>
      </td>
      <%if not isNull(eBullet) and eBullet<>"0" then%>
      <td width="10" valign="top" bgcolor="<%=eBGColor%>">     
		<img src="../../images/bul<%=eBullet%>.gif" border="0">      
	  </td> 	
      <%end if%>
      <%if (eSpace<>nil) then%>  
      <td width="<%=eSpace%>" nowrap></td>
      <%end if%>         
      </tr>
      </table>
    </td>
	<td width="1%" align="center" valign="bottom" bgcolor="#DDDDDD">
	<%if i>1 then%>
		<a href="event.asp?pageId=<%=pageId%>&prodId=<%=prodId%>&amp;up=1&amp;place=<%=eNumOnPage%>"><img src="images/up.gif" alt="להעביר את הטקסט למעלה" border="0" width="11"></a>
	<%else%>
	  &nbsp;
	<%end if%>
	</td>
</tr>
<tr>
	<td width="3%" align="center" valign="top" bgcolor="#DDDDDD">
   <a href="event.asp?pageId=<%=pageId%>&DEL=1&amp;elemId=<%=elemId%>&amp;prodId=<%=prodId%>&amp;place=<%=eNumOnPage%>" ONCLICK="return CheckDel()" class="but_delete">מחיקה</a>
	</td>
	<td width="1%" align="center" valign="top" bgcolor="#DDDDDD">
	<%if i<mass then%>
		<a href="event.asp?pageId=<%=pageId%>&prodId=<%=prodId%>&amp;down=1&amp;place=<%=eNumOnPage%>"><img src="images/down.gif" alt="להעביר את הטקסט למטה" border="0" width="11"></a>
	<%else%>
	  &nbsp;
	<%end if%>
	</td>
</tr>
<%
End Function

Function imageFunction(i,mass,elemId,eNumOnPage,eOrderAlign,eText,eLinkStatus,eLinkTarget,eLink,eBullet,eWWord,ePictureSize,eFontSize,eFontType,eFontColor,eBGColor,eHebrSize,eAlign,eTxtImgAlign,eHomePage,eConnection,eVis,ePicWidth,ePicHight)
	if not isNull(eText) then
		eTextH=eText
		pr=left(eTextH,1)
		'if eHebrSize<>nil then	
			'eTextH=con.hebrew(txt_for_hebr,eHebrSize)
		'else		
		  'koeff=120\eFontSize	
		   	'eTextH=con.hebrew(txt_for_hebr,koeff)
		'end if
	end if
if eTextH=nil or eTextH="" then
	eTextH=" "
end if
If Len(trim(eLinkTarget)) = 0 Then
	eLinkTarget = "_blank"
End If
if len(Trim(eWWord))>0 then		
		eWWordH=eWWord
		tmpStr="<font style=font-size:" & 6+eWSize*2 & "pt; color=" & eWColor & "><" & eWType & ">" & eWWordH & "</" & eWType & "></font>"
		eTextH=replace(eTextH,eWWordH,tmpStr)
end if	
%>
<tr>
	<td width="3%" bgcolor="#dddddd" valign="bottom" align="center">
	   <a href="AddPicture.asp?elemId=<%=elemId%>&amp;pageId=<%=pageId%>&prodId=<%=prodId%>" class="but_edit">עדכון</a>
	</td>	
	<td rowspan="2" align="<%=eAlign%>" valign="top" bgcolor="<%=eBGColor%>">
<table border="0" cellpadding="0" cellspacing="0" width="100%" align="<%=eAlign%>">
	<tr>
<%if eAlign="right" and eTextH<>nil then%>
	<%if pr="<" then%>
<td valign=middle width="100%" align="<%=eTxtImgAlign%>">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" align="<%=eTxtImgAlign%>">
		<tr>
		<td valign="top" dir=rtl style="padding-left : 5px; padding-right : 5px">
       <%=breaks(eText)%>
		</td>
       </tr>
    </table>
</td>
 <%   else%>
<td dir=rtl valign=top align="<%=eTxtImgAlign%>" style="padding-left : 5px; padding-right : 5px">
      <font color="<%=eFontColor%>" face="Ariel (Hebrew)" style="font-size:<%=6+eFontSize*2%>pt;">
	   <%if eFontType<>nil then%><%=chr(60)%><%=eFontType%><%=chr(62)%><%end if%>
		<%=breaks(eTextH)%>
       <%if eFontType<>nil then%><%=chr(60)%><%=chr(47)%><%=eFontType%><%=chr(62)%><%end if%>
      </font>
</td>
 <%   end if%>
<%end if%>
<td valign="top" align="<%=eAlign%>" width="10%">
		<%if ePictureSize>0 then %>
		<% If Not IsNull(ePicWidth) Then
			   If Not Trim(ePicWidth) = "0" Then
			   If Trim(ePicWidth) <> "" Then
			      res = res & " width='" &  ePicWidth & "'"
			   End If
			   End If
		   End If
                            
        If Not IsNull(ePicHight) Then
           If Not Trim(ePicHight) = "0" Then
           If Trim(ePicHight) <> "" Then
              res = res & " height='" & ePicHight & "'"
           End If
           End If
        End If
        %>
<%if eLinkStatus="global" then%><a href="<%=eLink%>" target=<%=eLinkTarget%>><%end if%><img src="../../GetImage.asp?DB=Template_Element&amp;FIELD=Element_Picture&amp;ID=<%=elemId%>" <%=res%> border="0" vspace="5" hspace="5"><%if eLinkStatus="global" then%></a><%end if%>
		<%else%>&nbsp;<%end if%>
</td>
<%if eAlign="left" and eTextH<>nil then%>
	<%if pr="<" then%>
<td valign=middle width="100%" align="<%=eTxtImgAlign%>">
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
		<tr>
		<td valign="top" dir=rtl style="padding-left : 5px; padding-right : 5px">
       <%=eText%>
		</td>
       </tr>
    </table>
</td>
 <%   else%>
 <td dir=rtl valign=top align="<%=eTxtImgAlign%>" style="padding-left : 5px; padding-right : 5px">
      <font color="<%=eFontColor%>" style="font-size:<%=6+eFontSize*2%>pt;">
	   <%if eFontType<>nil then%><%=chr(60)%><%=eFontType%><%=chr(62)%><%end if%>
		<%=breaks(eTextH)%>
       <%if eFontType<>nil then%><%=chr(60)%><%=chr(47)%><%=eFontType%><%=chr(62)%><%end if%>
      </font>
</td>
	<%end if%>
<%end if%>
	</tr>
<%if eAlign="center" and eTextH<>nil then%>	
	<%if pr="<" then%>
<tr>
	<td width="100%" valign=middle align="<%=eTxtImgAlign%>">
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
		<tr>
		<td valign=top dir=rtl style="padding-left : 5px; padding-right : 5px">
       <%=eText%>
		</td>
       </tr>
    </table>
	</td>
</tr>
 <%   else%>
	<tr>
	  <td dir=rtl valign=top align="<%=eTxtImgAlign%>" style="padding-left : 5px; padding-right : 5px">
      <font color="<%=eFontColor%>" style="font-size:<%=6+eFontSize*2%>pt;">
	   <%if eFontType<>nil then%><%=chr(60)%><%=eFontType%><%=chr(62)%><%end if%>
		<%=breaks(eTextH)%>
       <%if eFontType<>nil then%><%=chr(60)%><%=chr(47)%><%=eFontType%><%=chr(62)%><%end if%>
      </font>
		</td>
	</tr>
	<%end if%>
<%end if%>
</table>
    </td>
	<td width="1%" align="center" valign="bottom" bgcolor="#dddddd">
	<%if i>1 then%>
		<a href="event.asp?pageId=<%=pageId%>&prodId=<%=prodId%>&amp;up=1&amp;place=<%=eNumOnPage%>"><img src="images/up.gif" alt="להעביר את התמונה למעלה" border="0" width="11"></a>
	<%else%>
	  &nbsp;
	<%end if%>
	</td>
</tr>
<tr>
	<td width="3%" align="center" valign="top" bgcolor="#dddddd">
<a href="event.asp?pageId=<%=pageId%>&DEL=1&amp;elemId=<%=elemId%>&amp;prodId=<%=pageId%>&amp;place=<%=eNumOnPage%>" ONCLICK="return CheckDel()" class="but_delete">מחיקה</a>
	</td>
	<td width="1%" align="center" valign="top" bgcolor="#dddddd">
	<%if i<mass then%>
		<a href="event.asp?pageId=<%=pageId%>&prodId=<%=prodId%>&amp;down=1&amp;place=<%=eNumOnPage%>"><img src="images/down.gif" alt="להעביר את התמונה למטה" border="0" width="11"></a>
	<%else%>
	  &nbsp;
	<%end if%>
	</td>
</tr>
<%
End Function
%>


<%
Function HTMLFunction(i,mass,elemId,eNumOnPage,eOrderAlign,eText,eFileName,eFontSize,eFontType,eFontColor,eBGColor,eHebrSize,eAlign,eHomePage,eConnection,eVis)
        if isNull(eText) then
           eText=" "
         end if  
        If trim(eBGColor) = "" Or isNULL(eBGColor) Then
			eBGColor="transparent"
		End If	 
		
%>
<tr>
	<td width="3%" align="center" valign="bottom" bgcolor="#dddddd">
	<a href="editHTML.asp?elemId=<%=elemId%>&pageId=<%=pageId%>&prodId=<%=prodId%>" class="but_delete">עדכון</a></td>
	<td rowspan="2" align="<%=center%>" valign="top" bgcolor="<%=eBGColor%>">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
		<tr><td>		
       <%=eText%>		
       </td>
       </tr>
       </table>
    </td>
	<td width="1%" align="center" valign="bottom" bgcolor="#dddddd">
	<%if i>1 then%>
		<a href="event.asp?pageId=<%=pageId%>&prodId=<%=prodId%>&amp;up=1&amp;place=<%=eNumOnPage%>"><img src="images/up.gif" alt="להעביר את הקובץ למעלה" border="0" width="11"></a>
	<%else%>
	  &nbsp;
	<%end if%>
	</td>
</tr>
<tr>
	<td width="3%" align="center" valign="top" bgcolor="#dddddd">
<a href="event.asp?pageId=<%=pageId%>&DEL=1&amp;elemId=<%=elemId%>&amp;prodId=<%=prodId%>&amp;place=<%=eNumOnPage%>" ONCLICK="return CheckDel()" class="but_delete">מחיקה</a>
	</td>
	<td width="1%" align="center" valign="top" bgcolor="#dddddd">
	<%if i<mass then%>
		<a href="event.asp?pageId=<%=pageId%>&prodId=<%=prodId%>&amp;down=1&amp;place=<%=eNumOnPage%>"><img src="images/down.gif" alt="להעביר את הקובץ למטה" border="0" width="11"></a>
	<%else%>
	  &nbsp;
	<%end if%>
	</td>
</tr>
<%
End Function
%>