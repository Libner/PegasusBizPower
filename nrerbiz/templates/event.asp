<!--#include file="..\..\netcom/connect.asp"-->
<!--#INCLUDE file="..\..\netcom/reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%
  prodId=Request("prodId")
  pageId=Request("pageId")
  elemId=Request("elemId")

  down=Request.QueryString("down")
  up=Request.QueryString("up")
  place=Request("place")
  newPlace=CInt(place)+1
  
  isLink=Request.Form("isLink")
  bull=Request.Form("bullet")
  if bull=nil then 
	bull="0"
  end if
  
if isLink=nil then
	isLink="noLink"
else
	isLink=Trim(isLink)
end if


if pageId<>nil then
sql="select count(*) as myCont from Template_Elements where Page_Id="&pageId&" "
'Response.Write sql
'Response.End
set mcnt=con.getRecordSet(sql)
	cnt=mcnt("myCont")
mcnt.close


set myNum=con.getRecordSet("Select Element_Num_On_Page,Element_Id,Page_Id from Template_Elements where Page_Id="&pageId&" order by Element_Num_On_Page ")

if not myNum.eof then
myNum.moveFirst
i=1
do while i<=cnt
elId=myNum("Element_Id")
  stst="update Template_Elements set Element_Num_On_Page="&i&" where Element_Id=" & elId & " "
  con.getRecordSet(stst)
  'Response.Write "<br>stst="&stst
myNum.moveNext
i=i+1
loop
end if

myNum.close
End If 


if request.form("editTxt")="1" then		

 if elemId<>nil then
	 
	upEx="update Template_Elements set Element_Text='"&sFix(Request.form("title1"))&"',Element_BGColor='"&sFix(Request.form("bgcolor"))&"',Element_BGColor_Name='"&sFix(Request.form("bgcolname"))&"',Element_Bullet="&bull&",Element_Word='"&sFix(Request.form("WWord"))&"',Elemen_Word_Color='"&Request.form("wColor")&"',Element_Word_Color_Name='"&Request.form("wColname")&"',Element_Word_Size="&Request.form("wSize")&",Element_Word_Type='"&Request.form("wStyle")&"',Element_Link_Status='"&isLink&"',Element_Link_Target='"&trim(request("linkTarget"))&"',Element_Font_Size="&Request.form("fsize")&",Element_Font_Type='"&Request.form("fstyle")&"',Element_Font_Color='"&Request.form("fcolor")&"',Element_Font_Color_Name='"&Request.form("fcolname")&"',Element_Align='"&Request.form("falign")&"',Element_Space="&Request.Form("tSpace")&" where Element_ID="&elemId&" "
	'Response.Write("<brupEx=>"&upEx)	
	'Response.End
	con.getRecordSet(upEx) 	 
	con.getRecordSet "delete from Template_Link_Texts WHERE Element_Id="&elemId&" "
	 else
		if newPlace>0 then
		'Response.Write "UPDATE Template_Elements set Element_Num_On_Page=Element_Num_On_Page+1 WHERE Element_Num_On_Page>="&newPlace&" and Page_Id="&pageId&" "
		'Response.End
			con.getRecordSet "UPDATE Template_Elements set Element_Num_On_Page=Element_Num_On_Page+1 WHERE Element_Num_On_Page>="&newPlace&" and Page_Id="&pageId&" "
		end if
		sqlstring="INSERT INTO Template_Elements(Page_ID,Element_Num_On_Page,Element_Type,Element_Text,Element_Bullet,Element_Word,Elemen_Word_Color,Element_Word_Color_Name,Element_Word_Size,Element_Word_Type,Element_Link_Status,Element_Link_Target,Element_Font_Size,Element_Font_Type,Element_Font_Color,Element_Font_Color_Name,Element_Align,Element_Space,Element_BGColor,Element_BGColor_Name) VALUES ("&pageId&","&newPlace&",1,'"&sFix(Request.form("title1"))&"',"&bull&",'"&sFix(Request.form("WWord"))&"','"&Request.form("wColor")&"','"&Request.form("wColname")&"',"&Request.form("wSize")&",'"&Request.form("wStyle")&"','"&isLink&"','"&trim(request("linkTarget"))&"',"&Request.form("fsize")&",'"&Request.form("fstyle")&"','"&Request.form("fcolor")&"','"&Request.form("fcolname")&"','"&Request.form("falign")&"',"&Request.Form("tSpace")&",'"&sFix(Request.form("bgcolor"))&"','"& sFix(Request.form("bgcolname"))& "') "
		'Response.Write "<br>sqlstring="&sqlstring
		'Response.End
		con.getRecordSet(sqlstring)
		
		set elem=con.getRecordSet("SELECT * from Template_Elements ORDER BY Element_ID DESC")
		    elemId=CStr(elem("Element_ID"))
		elem.close
	end if
		if isLink="global" then
		 if Request.form("allLink")<>nil then
			if Request.form("allColor")<>nil then
				allCol=Request.form("allColor")
			else
				allCol="#0000ff"
			end if
			  con.getRecordSet "insert into Template_Link_Texts(Element_Id,Key_Text,Link_Text,Link_Color) VALUES ("&elemId&",'"&sFix(Request.form("title1"))&"','"&Request.form("allLink")&"','"&allCol&"') "
			 tup="update Template_Elements set Element_Link='"&Request.form("allLink")&"' WHERE Element_ID="&elemId&" "
			  con.getRecordSet (tup)
		 else
			  con.getRecordSet "UPDATE Template_Elements SET Element_Link_Status='noLink' WHERE Element_Id="&elemId	  
		 end if	
		end if

		if isLink="words" then
		 if ( Request.form("wordLink1")<>nil and Request.form("word1")<>nil ) then
			if Request.form("wordColor1")<>nil then
				wordCol1=Request.form("wordColor1")
			else
				wordCol1="#0000ff"
			end if
			  con.getRecordSet "insert into Template_Link_Texts(Element_Id,Key_Text,Link_Text,Link_Color) VALUES ("&elemId&",'"&sFix(Request.form("word1"))&"','"&sFix(Request.form("wordLink1"))&"','"&wordCol1&"') "
		 end if	
		end if

end if

if ( request.form("editPicture")="1" and elemId<>nil ) then
    Dim width_pic	
	Dim height_pic
	
	width_pic = trim(Request.Form("width_pic"))
	If IsNumeric(width_pic) = false Then width_pic = "NULL"	
	height_pic = trim(Request.Form("height_pic"))
	If IsNumeric(height_pic) = false Then height_pic = "NULL"
	
  if request.form("pictureHTML")="1" then
'	Response.Write "numcolumn="& request("numcolumn")
	if request("numcolumn")="on" and request("htmltext2")<>nil then
		n=Fix((45-3*request("columnnumb"))/request("columnnumb"))
'		Response.Write ("n="&n&", "&45/request("columnnumb"))
		htmltext=reverseHTML(request("htmltext2"),n)
		htmltext=Replace(htmltext,"&lt;","<")
		htmltext=Replace(htmltext,"&gt;",">")
		
	else
		htmltext=(request.Form("title1"))
	end if
'		Response.Write("numcolumn="&request("numcolumn")&" title1="&Request.Form("title1")&" htmltext2="&request("htmltext2")&" htmltext="&htmltext)

   
	 if elemId<>nil then
		upEx="update Template_Elements set Element_Text='"&sFix(htmltext)&"' ,Element_Link_Status='"&isLink&_
		"',Element_Link_Target='"&trim(request("linkTarget"))&"',Element_Link='"&sFix(Request.form("allLink"))&_
		"',Element_Align='"&Request.form("falign")&"',Element_BGColor='"&sFix(Request.form("bgcolor"))&_
		"',Element_BGColor_Name='"&sFix(Request.form("bgcolname"))&"',Element_Bullet="&bull&_
		",Element_Word='"&sFix(Request.form("WWord"))&"',Elemen_Word_Color='"&Request.form("wColor")&_
		"',Element_Word_Color_Name='"&Request.form("wColname")&_
		"',Element_Word_Size="&Request.form("wSize")&",Element_Word_Type='"&Request.form("wStyle")&_
		"',Element_pic_width= " & width_pic & ",Element_pic_height= " & height_pic &_
		"  where Element_ID="&elemId&" "
		'Response.Write(upEx)	
		'Response.End
		con.getRecordSet(upEx) 	 
	 else
		if newPlace>0 then
			con.getRecordSet "UPDATE Template_Elements set Element_Num_On_Page=Element_Num_On_Page+1 WHERE Element_Num_On_Page>="&newPlace&" and Page_Id="&pageId&" "
		end if
		sqlstring="INSERT INTO Template_Elements(Page_ID,Element_Num_On_Page,Element_Type,Element_Text,Element_Space,Element_BGColor,Element_BGColor_Name,Element_Bullet,Element_Word,Elemen_Word_Color,Element_Word_Color_Name,Element_Word_Size,Element_Word_Type,Element_pic_width,Element_pic_height)"&_
		" VALUES ("&prodId&","&newPlace&",5,'"&sFix(htmltext)&"',0,'"&sFix(Request.form("bgcolor"))&"','"& sFix(Request.form("bgcolname"))&",'"&sFix(Request.form("WWord"))&"','"&Request.form("wColor")&"','"&Request.form("wColname")&"',"&Request.form("wSize")&",'"&Request.form("wStyle")& "'," & width_pic & "," & height_pic & ")"
		'Response.Write(sqlstring)
		con.getRecordSet(sqlstring)
			set elem=con.getRecordSet("SELECT * from Template_Elements ORDER BY Element_ID DESC")
			    elemId=CStr(elem("Element_ID"))
			elem.close
	end if
  else	'if request.form("pictureHTML")="1"
        width_pic = trim(Request.Form("width_pic"))
		If IsNumeric(width_pic) = false Then width_pic = "NULL"	
		height_pic = trim(Request.Form("height_pic"))
		If IsNumeric(height_pic) = false Then height_pic = "NULL"
		sqlUpdt="update Template_Elements set Element_Text='"&sFix(Request.form("title1"))&"',Element_Link_Status='"&isLink&"',Element_Link_Target='"&trim(request("linkTarget"))&"',Element_Link='"&sFix(Request.form("allLink"))&"',Element_Font_Size='"&Request.form("fsize")&"',Element_Font_Type='"&Request.form("fstyle")&"',Element_Font_Color='"&Request.form("fcolor")&"',Element_Font_Color_Name='"&Request.form("fcolname")&"',Element_Align='"&Request.form("falign")&"',Element_Text_Image_Align='"&Request.form("fTxtAlign")&"',Element_BGColor='"&sFix(Request.form("bgcolor"))&"',Element_BGColor_Name='"&sFix(Request.form("bgcolname"))&"',Element_Bullet="&bull&",Element_Word='"&sFix(Request.form("WWord"))&"',Elemen_Word_Color='"&Request.form("wColor")&"',Element_Word_Color_Name='"&Request.form("wColname")&"',Element_Word_Size="&Request.form("wSize")&",Element_Word_Type='"&Request.form("wStyle")&_
		"',Element_pic_width= " & width_pic & ",Element_pic_height= " & height_pic &_
		" where Element_ID="&elemId&" "
		'Response.Write sqlUpdt
		'Response.End
		con.getRecordSet(sqlUpdt)  
  end if
end if 

if request.form("editHTML")="1" then
	htmltext=trim(request("htmltext2"))
	if request("numcolumn")="2" then   
	n=Fix((95-(request("columnnumb")^2))/request("columnnumb"))	
	htmltext=Replace(htmltext,"&lt;","<")
	htmltext=Replace(htmltext,"&gt;",">")
	else
		htmltext=(request("htmltext"))
		htmltext=Replace(htmltext,"----------","")
		htmltext=Replace(htmltext,"כתוב תוכן של התא","")
		htmltext=Replace(htmltext,"שורה חדשה","")		
	end if
	 if elemId<>nil then
		upEx="update Template_Elements set Element_Text='"&sFix(htmltext)&"' where Element_ID="&elemId&" "
			'Response.Write(upEx)	
		con.getRecordSet(upEx) 	 
	 else
		if newPlace>0 then
			con.getRecordSet "UPDATE Template_Elements set Element_Num_On_Page=Element_Num_On_Page+1 WHERE Element_Num_On_Page>="&newPlace&" and Page_Id="&pageId&" "
		end if
		sqlstring="INSERT INTO Template_Elements(Page_ID,Element_Num_On_Page,Element_Type,Element_Text,Element_Space) VALUES ("&pageId&","&newPlace&",5,'"&sFix(htmltext)&"',0) "
		'Response.Write(sqlstring)
		'Response.End
		con.getRecordSet(sqlstring)
			set elem=con.getRecordSet("SELECT * from Template_Elements ORDER BY Element_ID DESC")
			    elemId=CStr(elem("Element_ID"))
			elem.close
	end if
	
end if

if Request.QueryString("DEL")="1" then 
	set fileName=con.getRecordSet("SELECT Element_File_Name,Element_Type from Template_Elements where Element_File_Name is not Null and Element_Id="&elemId ) 
	if not fileName.eof then
	  if fileName("Element_Type")<>4 then
	set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
		fileString= Server.MapPath("../../Download_h/"&fileName("Element_File_name")) 'server.MapPath is the current path we're in
		if fs.fileExists(fileString) then
			set f=fs.GetFile(fileString) 
			f.Delete TRUE   'TRUE means: delete even if the file is Read-Only 
	    end if
	  end if	
	end if
	con.getRecordSet "delete from Template_Link_Texts where Element_ID="&elemId&" "
	con.getRecordSet "delete from Template_Elements where Element_ID="&elemId&" "
	con.getRecordSet "update Template_Elements set Element_Num_On_Page=Element_Num_On_Page-1 WHERE Element_Num_On_Page>"&place&" and Page_Id="&pageId&" "
	Response.Redirect ("event.asp?prodId="&prodId&"&pageId="&pageId)
end if	

if Request.QueryString("vis")="on" then 
	con.getRecordSet "update Template_Elements set Element_MyVisible=1 WHERE Element_Id="&ElemId&" "
end if	

if Request.QueryString("vis")="off" then 
	con.getRecordSet "update Template_Elements set Element_MyVisible=0 WHERE Element_Id="&ElemId&" "
end if	

if down<>nil then
stam=place+1
	con.getRecordSet("update Template_Elements set Element_Num_On_Page=-10 WHERE Element_Num_On_Page="&stam&" and Page_Id="&pageId&" ")
	con.getRecordSet("update Template_Elements set Element_Num_On_Page="&stam&" WHERE Element_Num_On_Page="&place&" and Page_Id="&pageId&" ")
	con.getRecordSet("update Template_Elements set Element_Num_On_Page="&place&" WHERE Element_Num_On_Page=-10 ")
end if

if up<>nil then
stam=place-1
	con.getRecordSet("update Template_Elements set Element_Num_On_Page=-10 WHERE Element_Num_On_Page="&stam&" and Page_Id="&pageId&" ")
	con.getRecordSet("update Template_Elements set Element_Num_On_Page="&stam&" WHERE Element_Num_On_Page="&place&" and Page_Id="&pageId&" ")
	con.getRecordSet("update Template_Elements set Element_Num_On_Page="&place&" WHERE Element_Num_On_Page=-10 ")
end if

If pageId<>nil Then
set countTest=con.getRecordSet("SELECT count(*) as cntTest FROM Template_Elements WHERE Page_Id="&pageId&" ")
	mass=countTest("cntTest") ' for DownButton and UpButton
countTest.close	
End If
set typesLinks=con.getRecordSet("SELECT * FROM Types ")
sqlpg="SELECT * FROM Templates WHERE Page_Id="&pageId&""
'Response.Write sqlpg
'Response.End
set pg=con.getRecordSet(sqlpg)
perSize1=pg.Fields("Page_Picture").ActualSize
pKeyWords=pg("Page_Keywords")
percolor=pg("Page_Bgcolor")
pertitle=pg("Page_Title")
peralign=pg("Page_Align")
perwidth=pg("Page_Width")
Pagebgcolorname=pg("Page_Bgcolor_Name")

%>
<html>
<head>
<title>ניהול תבניות דפי מבצע</title>
<meta charset="windows-1255">
<link href="../../admin.css" rel="STYLESHEET" type="text/css">

<SCRIPT LANGUAGE=javascript>
<!--
	function openPreview()
	{
		
		url="result.asp?prodId=<%=prodId%>&pageId=<%=pageId%>";
		window.open(url,"_blank","width=800,left=0,status=no,toolbar=no,menubar=no,location=no,scrollbars=yes");
	}
	
	function CheckDel()
	{
		return window.confirm("?האם ברצונך למחוק את האלמנט");
	}
	
//-->
</SCRIPT>

</head>
<!--#INCLUDE FILE="functions.asp"-->

<body topmargin=0 leftmargin=0 rightmargin=0 bottommargin=0>
<table bgcolor="#c0c0c0" border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="5" class="page_title">&nbsp;עריכת&nbsp;תבנית</td>
  </tr>
  <tr>
     <td width="5%" align="center"><a class="button_admin_1" onClick="return openPreview()" href="#">תצוגה מקדימה</a></td>  
     <td width="5%" align="center"><a class="button_admin_1" href="default.asp">חזרה לדפי מבצע</a></td>     
     <td width="5%" align="center"><a class="button_admin_1" href="../choose.asp">חזרה לדף ניהול ראשי</a></td>  
     <td width="*%" align="center"></td> 
  </tr>
</table>
<table border="0" width="100%"  cellspacing="1" cellpadding="0">
	<tr>
		<td align="center"colspan=2 width=100%>&nbsp;</td>
			</tr>
</table>
<table width="780" cellspacing="1" cellpadding="0" align=center border="0" bgcolor="#ffffff" >

<tr><td height=10 nowrap></td></tr>
<tr><td width=100%><table width="685" border=0 cellpadding=0 cellspacing=0 align="center">
<tr>
   <td nowrap width=685 align="right" bgcolor="#DDDDDD" colspan="8">	       
    <table border="0" cellpadding="0" cellspacing="0" width="685">
     <tr>
     <td width=60%></td>
<%do while not typesLinks.Eof%>
  <td align="center" nowrap>
    &nbsp;<a href="<%=typesLinks("Type_Add_File_Name")%>.asp?pageId=<%=pageId%>&prodId=<%=prodId%>&amp;place=<%=0%>"><img src="images/<%=typesLinks("Type_Image")%>" alt="<%=typesLinks("Type_Alt")%>" border="0"></a>
  </td>
<%typesLinks.movenext
  Loop%>			      
      </tr>
     </table>
    </td>		   
   </tr>

<%
If isNumeric(trim(pageId)) = true Then
i=1 ' for DownButton and UpButton
sqlSt="SELECT * FROM Template_Elements WHERE Page_Id="&pageId&" Order By Element_Num_On_Page"
set el=con.getRecordSet(sqlSt)
 do while not el.eof

pageId=el("Page_Id")
elemId=el("Element_Id")
eNumOnPage=el("Element_Num_On_Page")
eOrderAlign=el("Element_Order_Align")
eType=el("Element_Type")
eText=el("Element_Text")
eLinkStatus=Trim(el("Element_Link_Status"))
eLink=el("Element_Link")
eLinkTarget=Trim(el("Element_Link_Target"))

eBullet=el("Element_Bullet")
eWWord=el("Element_Word")
eWColor=el("Elemen_Word_Color")
eWColorName=el("Element_Word_Color_Name")
eWSize=el("Element_Word_Size")
eWType=el("Element_Word_Type")

eBGColor = el("Element_BGColor")
If trim(eBGColor) = "" And trim(percolor) <> "" Then
	eBGColor = percolor
ElseIf trim(eBGColor) = "" Or perSize1>0 Then
   eBGColor = "transparent"	
Else
	eBGColor = percolor   
End If

eFileName=el("Element_File_Name")
ePictureSize=el.Fields("Element_Picture").ActualSize
eFontSize=el("Element_Font_Size")
eFontType=el("Element_Font_Type")
eFontColor=el("Element_Font_Color")
eHebrSize=el("Element_Hebrew_Size")
eAlign=Trim(el("Element_Align"))
eTxtImgAlign=Trim(el("Element_Text_Image_Align"))
eSpace=el("Element_Space")
ePicWidth = el("Element_pic_width")
ePicHight = el("Element_pic_height")

%>
<%
CheckType i,mass,eType,elemId,eNumOnPage,eOrderAlign,eText,eLinkStatus,eLinkTarget,eLink,eBullet,eWWord,eWColor,eWColorName,eWSize,eWType,eFileName,ePictureSize,eFontSize,eFontType,eFontColor,eBGColor,eHebrSize,eAlign,eHomePage,eConnection,eVis,ePicWidth,ePicHight
i=i+1 ' for DownButton and UpButton
typesLinks.movefirst
%>

<tr bgcolor="#DDDDDD">
	<td colspan="6" width="100%" nowrap>
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
		  <tr>	      
	       <td nowrap width=685 align="right">
	        <table border="0" align="right" cellpadding="0" cellspacing="0" width="685">
	         <tr>
	          <td width="60%" nowrap>&nbsp;</td>
		<%do while not typesLinks.Eof%>
		  <td align="center">
		    &nbsp;<a href="<%=typesLinks("Type_Add_File_Name")%>.asp?pageId=<%=pageId%>&prodId=<%=prodId%>&amp;place=<%=eNumOnPage%>"><img src="images/<%=typesLinks("Type_Image")%>" alt="<%=typesLinks("Type_Alt")%>" border="0"></a>
		  </td>
		<%typesLinks.movenext
		  Loop%>	
		      </tr>
		     </table>
		    </td>		  
		   </tr>
		</table>
	</td>
</tr>
<%
el.movenext
Loop
End If
%>
</table>
</td></tr></table>
</body>
</html>
