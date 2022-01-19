<!--#include file="../connect.asp"-->
<%
	pageId = trim(Request("pageId"))
	templateId = trim(Request("templateId"))
	perSize2 = trim(Request("perSize2"))
	
	If isNumeric(perSize2) = True Then
		perSize2 = cLng(perSize2)
	Else
		perSize2 = 0
	End If		
%>
<table background="<%=strLocal%>/netcom/templates/images/top5&#46;gif" border="0" width="620" height="76" cellspacing="0" cellpadding="0">
  <tr>
    <td width="620" valign="top">&nbsp;
     <%If pageId<>"" And perSize2 > 0 Then%>
    <img id="imgPict" name="imgPict" src="<%=strLocal%>/netcom/GetImage.asp?DB=Page&amp;FIELD=Organization_Logo&amp;ID=<%=pageId%>" border="0" hspace=5 vspace=5>
    <%End If%>
    <%If templateId<>"" And perSize2 > 0 Then%>
    <img id="imgPict" name="imgPict" src="<%=strLocal%>/netcom/GetImage.asp?DB=Template&amp;FIELD=Organization_Logo&amp;ID=<%=templateId%>" border="0" hspace=5 vspace=5>
    <%End If%>
    </td>
  </tr>
</table>

