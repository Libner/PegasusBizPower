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
    
    childOverlap = 548;
    childOffset = 0;
    perCentOver = null;
    secondsVisible = .5;
    menuWidth = 120; 
    fntCol = "#FFFFFF";  
    fntSiz = "9";
    fntBold = false;      
    fntItal = false;         
    fntFam = "arial";    
    backColor = "#ffd012";   
    overColor = "#3a3184";
    overFntColor = "#6E6DA6"; 
    borWid = 0;
    borderCol = "#ffd012";
    borSty = "solid";        
    itemPad = 3;             
    imgSrc = "../images/triangle.gif";
    imgSiz = 10;        
    separator = 1;           
    separatorColor = "transparent";     
    isFrames = false;         
    navFrLoc =  "right";    
    mainFrName ="main";   
    keepHilite = true;   
    NSfontOver = true;   
    clickStart = false;   
    clickKill = false;   
    showVisited = false;
    menuVersion = 3;
}
<%
sub category_links(catid)
	for i=0 to ubound(subcatname,2)
		if subcatname(catid-1,i)<>nil then
			If trim(catid) = "6" Then
				is_links = 1				
			Else
				is_links = 0
			End If		
        %>
			,"<%=subcatname(catid-1,i)%>&nbsp;&nbsp;","<%=subcatlink(catid-1,i)%>",<%=is_links%>	            
		<%end if
	next	
end sub
%>

<%
function category_sub_links(catid) 
	for i=0 to ubound(subcatname,2)
		if subcatname(catid-1,i)<>nil then			
        %>
			<%If i > 0 Then%>
			,
			<%End If%>
			"<%=subcatname(catid-1,i)%>&nbsp;","<%=subcatlink(catid-1,i)%>",0			
		<%end if		
	next	
end function
%>

/* Menu for links font-size:9pt:*/

	arMenu1 = new Array(90,5,94,9,'#ffd012','#dbdbf1','#3a3184','',''
	<%call category_links(1)%>
);

	arMenu2 = new Array(90,85,94,9,'#ffd012','#dbdbf1','#3a3184','',''
	<%call category_links(2)%>
);
	
	arMenu3 = new Array(90,120,94,9,'#ffd012','#dbdbf1','#3a3184','',''
	<%call category_links(3)%>
);
	
    arMenu4 = new Array(120,210,94,9,'#ffd012','#dbdbf1','#3a3184','',''
	<%call category_links(4)%>
);
	
	arMenu5 = new Array(120,255,94,9,'#ffd012','#dbdbf1','#3a3184','',''
	<%call category_links(5)%>
);


arMenu6 = new Array(100,345,94,9,'#ffd012','#dbdbf1','#3a3184','',''
	<%call category_links(6)%>
);

arMenu6_1 = new Array(<%=category_sub_links(7)%>);

arMenu6_2 = new Array(<%=category_sub_links(8)%>);

arMenu6_3 = new Array(<%=category_sub_links(9)%>);

arMenu6_4 = new Array(<%=category_sub_links(10)%>);

arMenu6_5 = new Array(<%=category_sub_links(11)%>);
</script>
<script LANGUAGE="JavaScript1.2" SRC="../PopUpMenus/PopUpMenus.js"></script>