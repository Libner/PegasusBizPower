<!--#INCLUDE FILE="../../include/reverse.asp"-->
<!--#include file="../../include/connect.asp"-->
<!--#INCLUDE FILE="../checkAuWorker.asp"-->
<html>
<head>
<title>Administration Site</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
<link rel="stylesheet" type="text/css" href="../admin_style.css">
<script LANGUAGE="JavaScript">
function CheckDelCat() {
  return (confirm("?האם ברצונך למחוק את הקטגוריה"))    
}
<!--End-->
</script>  
</head>
<body class="body_admin">
<div align="right">
<%
	dim MyPrCnt,cntCat,pl_Cat
    catId=request("catId")
    catname=request("catname")
        if catname<>nil and catname<>"" then
		if catId<>"" and catid<>nil then
			con.execute("UPDATE News_Categories SET Category_Name='" & catname & "' WHERE Category_ID=" & catid)
		else
			ord=0
			set ordcat=con.execute("SELECT Category_Order AS maxord FROM News_Categories ORDER BY Category_Order DESC")
			if not ordcat.EOF then
			   ord=ordcat("maxord")
			end if   
			con.execute("INSERT INTO News_Categories (Category_Name,Category_Order) VALUES ('" & catname & "'," & ord+1 & ")")
		end if
    end if
    delcatId=request("delcatId")
    if delcatId<>nil and delcatId<>"" then
		s="delete from News_Categories WHERE Category_ID=" & delcatid
		d=0
		con.execute s,d
		if d<>0 then
		set pr=con.execute("SELECT New_Picture FROM News WHERE Category_Id=" & delcatid)
		do while not pr.eof
		fileName1=pr("New_Picture")
		if fileName1<>"" then
			set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
			fileString= Server.MapPath("../../download/pictures/"& fileName1 ) 'deleting of existing file
			if fs.FileExists(fileString) then
				set f=fs.GetFile(fileString) 
				f.Delete TRUE   'TRUE means: delete even if the file is Read-Only 
			end if
	    end if
	    pr.movenext	
		loop
		pr.close
		s="delete from News where Category_Id=" & delcatid
		con.execute (s)
		end if
	end if
  %>
<table border="0" width="100%" align="center"  cellspacing="0" cellpadding="0">
  <tr><td class="a_title_big" width="100%" dir=rtl><font size="4" color="#ffffff"><strong>ניהול אירועים</strong></font></td></tr>
  <tr>
    <td>
       <table class="table_admin_1" border="0" width="100%" align="center"  cellspacing="0" cellpadding="0">
         <tr>
             <td class="td_admin_5" align="left" nowrap><a class="button_admin_1" href="../choose.asp">חזרה לדף ניהול ראשי</a></td>
             <td class="" align="right" valign="middle" width="100%"><font style="font-size:12pt;"><b>ניהול קטגוריות</b>&nbsp;</font></td>
         </tr>
       </table>
    </td>
  </tr>
  <tr><td height="5"></td></tr>

</table>
        <table width="100%" align="center" border="0" cellpadding="2" cellspacing="1">
          <%set_Orders_Categories
          ' rem ==============================
          sss="SELECT Category_ID,Category_Order,Category_Name FROM News_Categories  ORDER BY Category_Order"
          set pr=con.execute(sss)
	      pl_Cat=1
	      do while not pr.EOF
		     catId=pr("Category_ID")
		     cat_ord=pr("Category_Order")
		     cat=pr("Category_Name")
		     set ord=con.Execute("select count(New_ID) as prNum from News WHERE New_Site_Visible=1 AND Category_Id=" & catId)
                 MyPrCnt=0
                 if not ord.EOF then
                    MyPrCnt=ord("prNum")
                 end if   
                 ord.close
                 set ord=Nothing%>
		    <tr>
				<%if 0 then%>
			   <td align="center" valign="top" class="td_admin_4"><a class="button_delete_1" href="adminCat.asp?delcatId=<%=catId%>" ONCLICK="return CheckDelCat()"><b>מחיקה</b></a></td>
			   <td align="center" valign="top" class="td_admin_4"><a class="button_edit_1" href="Acatadd.asp?catId=<%=catId%>"><b>עדכון</b></a></td>
               <%end if%>
               <td align="right" class="td_admin_4" width="80%"><a class="link_categ" href="admin.asp?catID=<%=catId%>"><%=cat%>&nbsp;</a></td>
               <%if 0 then%>
               <td width="2%" class="td_admin_2" valign="top" align="center">
                  <%if pl_Cat > 1 then%><a href="adminCat.asp?upCat=1&amp;placeCat=<%=cat_ord%>&amp;catId=<%=catId%>"><img src="../images/up.gif" border="0" WIDTH="11" HEIGHT="15"><%else%>&nbsp;<%end if%>
               </td>
               <td width="2%" class="td_admin_2" valign="top" align="center">
                  <%if pl_Cat < cntCat then%><a href="adminCat.asp?downCat=1&amp;placeCat=<%=cat_ord%>&amp;catId=<%=catId%>"><img src="../images/down.gif" border="0" WIDTH="11" HEIGHT="15"><%else%>&nbsp;<%end if%>
               </td>
               <%end if%>
               <td width="2%" class="td_admin_4" align="right" valign="top" nowrap>
                  <%=pl_Cat%>
               </td>
		    </tr>
				<%pl_Cat=pl_Cat+1
		        pr.movenext
		     loop		
	         pr.close 
             set pr=Nothing%>
             <tr><td colspan="6" align="center" bgcolor=#ffffff height="5" nowrap></td></tr>
          </table>
</div>

<%' REM: --------------------------------------------------------------
SUB set_Orders_Categories
' rem: Up & Down parents ================================
if Request.QueryString("downCat")<>nil then
	cat_Order=CInt(Request.QueryString("placeCat"))
	set ordCat=con.execute("select TOP 1 Category_Order FROM News_Categories WHERE Category_Order>"& cat_Order & " order by Category_Order ")
	if not ordCat.EOF then
	   nextord=ordCat("Category_Order")
       con.Execute("update News_Categories set Category_Order=-10 WHERE Category_Order="& nextord &" ")
       con.Execute("update News_Categories set Category_Order="& nextord &" WHERE Category_Order="& cat_Order &" " )
       con.Execute("update News_Categories set Category_Order="& cat_Order &" WHERE Category_Order=-10 " )
    end if
    set ordCat=Nothing
end if

if Request.QueryString("upCat")<>nil then
	cat_Order=CInt(Request.QueryString("placeCat"))
	set ordCat=con.execute("select TOP 1 Category_Order FROM News_Categories WHERE Category_Order<"& cat_Order  &" ORDER BY Category_Order desc")
	if not ordCat.EOF then
	   prevord=ordCat("Category_Order")
	   con.Execute("update News_Categories set Category_Order=-10 WHERE Category_Order="& prevord &" ")
       con.Execute("update News_Categories set Category_Order="& prevord &" WHERE Category_Order="& cat_Order &" ")
       con.Execute("update News_Categories set Category_Order="& cat_Order &" WHERE Category_Order=-10 ")
    end if
    set ordCat=Nothing
end if
cntCat=0
set mcntCat=con.Execute("select count(Category_ID) as myCont from News_Categories ")
if not mcntCat.EOF then
	cntCat=mcntCat("myCont")
end if
mcntCat.close
set myNumCat=con.Execute("select Category_ID,Category_Order FROM News_Categories ORDER BY Category_Order ")
pCat=1
do while pCat <= cntCat
pgId=myNumCat("Category_ID")
  stst="update News_Categories set Category_Order="& pCat &" where Category_ID=" & pgId
  con.Execute(stst)
myNumCat.moveNext
pCat=pCat + 1
loop
myNumCat.close
set myNumCat=Nothing
END SUB
%>
</body>
</html>
