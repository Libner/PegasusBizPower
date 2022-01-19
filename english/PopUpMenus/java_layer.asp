
<script>
   NS4 = (document.layers);
   IE4 = (document.all);
   ver4 = (NS4 || IE4);
   isMac = (navigator.appVersion.indexOf("Mac") != -1);
   isMenu = (NS4 || (IE4 && !isMac));
   function popUp(){return};
   function popDown(){return};
   function startIt(){return};
   if (!ver4) event = null;
</script>
<script>
if (isMenu) { 
    if (NS4){menuWidth = 140} else menuWidth = 140; 
    childOverlap = 0;
    childOffset = 0;
    perCentOver = null;
    secondsVisible = 0.2;
    fntCol = "#333333";  
    fntSiz = "9";
    fntBold = false;      
    fntItal = false;         
    fntFam = "arial";    
    backColor = "#ffd012";   
    overColor = "#3a3184";
    overFntColor = "#ffd012"; 
    borWid = 1;
    borderCol = "#C0C0C0";
    borSty = "solid";        
    itemPad = 3;             
    imgSrc = "";
    imgSiz = 10;     
    imgHeight = 7;               
    separator = 1;           
    separatorColor = "#C0C0C0";     
    isFrames = false;         
    navFrLoc =  "left";    
    mainFrName ="main";   
    keepHilite = true;   
    NSfontOver = true;   
    clickStart = false;   
    clickKill = false;   
    showVisited = false;
    menuVersion = 3;
}


/* Menu for links font-size:9pt:*/

<%
	sqlStr = "SELECT TOP 100 PERCENT Publication_Category_ID, Publication_Category_Name FROM dbo.Publication_Categories WHERE  Main_Category_Id=2 and  (Category_Vis = 1) AND (Publication_Category_ID IN  (SELECT     Category_ID FROM          PagesTav WHERE      (Page_Parent = 0) AND (Page_Visible_Title = 1))) ORDER BY Category_Order"
	set rs_Publication_Categories = con.execute(sqlStr)	
	i=1
	do while not rs_Publication_Categories.eof
		PubCategoryID = rs_Publication_Categories(0)
		PubCategoryName = rs_Publication_Categories(1)	
		if trim(PubCategoryName) <> "" then
%>
	arMenu<%=i%> = new Array(
	'','menuLayer<%=i%>',5,-10,'#616D66','#333333','#F3F3F3','#DDDDDD',''
		<%
		sqlStr = "SELECT TOP 100 PERCENT Page_Id, Page_Title, Page_URL FROM PagesTav WHERE (Page_Parent = 0) AND (Category_Id = "& PubCategoryID &") AND (Page_Visible_Title = 1) ORDER BY Page_Order"
		set rs_pages = con.execute(sqlStr)
		do while not rs_pages.eof
			currPageId = rs_pages(0)
			currPageTitle = rs_pages(1)	
			currPageURL = rs_pages(2)			
			
			if trim(currPageURL) = "" then
				href_string = "../template/default.asp?PageId="& currPageId & "&catId="& PubCategoryID & "&maincat=2"
			elseif instr(1,LCase(currPageURL),"http://")<>0 then
				href_string = "javascript:window.open('"& currPageURL &"')" ''currPageURL  ''" target='_new'"
			else
				href_string = currPageURL 
			end if								
			%>
			,"&nbsp;&nbsp;<%=vfix(currPageTitle)%>","<%=href_string%>",0	            
		<%
			rs_pages.movenext
			loop
			set rs_pages = nothing	
		%>
			
);
<%
		end if
	i=i+1		
	rs_Publication_Categories.movenext
	loop
     set rs_Publication_Categories = nothing
 %>



</script>
<!--script LANGUAGE="JavaScript1.2" SRC="../PopUpMenus/NavigationJScript.asp"></script-->
<script LANGUAGE="JavaScript1.2" SRC="../PopUpMenus/PopUpMenus.js"></script>
