   <script src="../javaScript/jquery-1.11.2.min.js"></script>	
    <script type="text/javascript" src="../javaScript/jquery.bpopup.min.js?v=<%=Application("sversion")%>"></script>			
<script>

function popupOpen(param_obj){
	var loadUrlType = "iframe";
	
	var href = param_obj.href;
    if (typeof (href) == "undefined") {
		href = "";
    }	
	var iframeAttr = param_obj.iframeAttr;
    if (typeof (iframeAttr) == "undefined") {
		iframeAttr = "";
    }	
	var post_data = param_obj.post_data;
    if (typeof (post_data) == "undefined") {
		post_data = false;
    }else{
		loadUrlType = "ajax";
    }	
        
    var customPopupNumber;
    customPopupNumber = "1";
    
    var myWindow;
    if (true || window.opener == null){
		myWindow = window;        
    }else{
		myWindow = window.opener;    		
    }
    
    
    if(myWindow.top.$('#customPopup1:visible').length > 0){
		closeBPopup();
		return false;
		 customPopupNumber = "2"
    }
        
    if(myWindow.top.$('#customPopup2:visible').length > 0){
		 customPopupNumber = "3"
    }
            
	myWindow.top.$('#customPopup' + customPopupNumber).bPopup({
			content: loadUrlType, //'ajax', 'iframe' or 'image'
			iframeAttr: iframeAttr,
			scrollBar: true,
			speed: 250,
			contentContainer:'.content' + customPopupNumber,
			modalClose: false,
			escClose: true,			           
			follow: [true,true], 
			followEasing: 'linear',
			followSpeed: 500,
			position: ['auto', 'auto'],
			amsl: 0,
			positionStyle: 'absolute',
			transition: 'fadeIn',
			transitionClose: 'fadeIn',
			loadData: post_data,
			loadUrl: href, 
			loadCallback: function(){ 
				//alert('loaded');
			}
		},
			function(){
				//alert('callback');
			}        		
		);
}

function closeBPopup() {
	//debugger;
    var myWindow;
    if (true || window.opener == null){
		myWindow = window;        
    }else{
		myWindow = window.opener;    		
    }
    	
	myWindow.$("#customPopup1").bPopup().close();
	myWindow.$("#customPopup1 .content1").empty();
location.reload();	
}
</script>



<%	Dim img_, bgr_pos, user_t, comp_t, show_logo, align_logo
	If trim(lang_id) = "2" Then
		  img_ = "_eng" : bgr_pos = "top right" : user_t = "User" : comp_t = "Company"
	Else
		  img_ = "" : bgr_pos = "top left" : user_t = "�����" : comp_t = "����"
	End If
   
	show_logo = true : align_logo = "center"
	If Not IsNull(Request.Cookies("bizpegasus")("UseBizLogo")) Then
		If trim(Request.Cookies("bizpegasus")("UseBizLogo")) = "0" Then 
			show_logo = false : align_logo = "left"
		End If	
	End If%>        
<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="ltr">
  <%  Set MessagesCount = Server.CreateObject("ADODB.RECORDSET")
   sqlstr = "exec dbo.get_NewMessages_Count " & Request.Cookies("bizpegasus")("UserId") 
   'Response.Write sqlstr
   'Response.End
   set  MessagesCount = con.getRecordSet(sqlstr)
     If not MessagesCount.EOF then		
        recCount = MessagesCount("NewMessages")		
   end if
   set MessagesCount=Nothing
      %>
   
  <%'if recCount>0 then%>
  <%'if false then%><tr><td><table border=0 cellpadding=0 cellspacing=0>
    <tr><td height=5></td></tr>
  <%if false then%>
       <tr><td><div id="customPopup1">
        <span class="button b-close"><span><i class="fa fa-times">X</i></span></span>
        <div class="content1">
        </div>
</div>
		</td></tr><%end if%>
		<%if false then%>
    <tr><td align=left bgcolor=#FF0000 style=color:#ffffff;padding-left:5px;padding-right:5px;>�� <%=recCount%> ������ �����</td></tr><%end if%>
    </table></td></tr><%'end if%>
    		<%'if false then%>
<tr><td><iframe id=checkMess src="http://bluto/bizpower_pegasus/NETcom/members/messages/checkNewMessagesWindow.aspx" frameborder="0" height="0"></iframe><%'end if%></td></tr>
<%'end if%>
    <tr>
      <td width="100%" valign="bottom" height="79">
      <table border="0" width="100%" style="height: 66px"  cellspacing="0" cellpadding="0">
        <tr>
          <td valign="top" align="center" width="100%" height="62"></td>
          <td valign="bottom" align="<%=align_var%>"><img src="<%=Application("VirDir")%>/images/topslog.gif" border="0" alt=""></td>
          <td width="201" valign=bottom>
          <table bgcolor="#6F6DA6" border="0" width="201" cellspacing="0" cellpadding="0" 
          style="background-image:URL('<%=Application("VirDir")%>/images/exit<%=img_%>.gif'); background-repeat:no-repeat;  background-position:top left;">
            <tr>
              <td width="201" nowrap height="21"><a href="<%=Application("VirDir")%>/exit.asp" target="_self"><img 
              src="<%=Application("VirDir")%>/images/top_knisa<%=img_%>.gif" border="0"></a></td>
            </tr>
            <tr>            
              <td width="100%" height="45">
              <table border="0" width="201" style="height: 45px" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">
                <tr>                                   
                  <td width=123 nowrap align="<%=align_var%>" dir="<%=dir_var%>" style="color:#FFD011;font-weight:600;vertical-align:middle;padding-<%=align_var%>:10px"><%=trim(user_name)%></td>
                  <td width=74 nowrap align="<%=align_var%>" dir="<%=dir_var%>" style="color:#FFFFFF;font-weight:600;vertical-align:middle;padding-<%=align_var%>:10px"><%=user_t%></td>
                </tr>
                <tr>
                 <td width=123 nowrap valign=top align="<%=align_var%>" dir="<%=dir_var%>" style="color:#FFD011;font-weight:600;vertical-align:middle;padding-<%=align_var%>:10px"><%=trim(org_name)%></td>
                 <td width=74 nowrap align="<%=align_var%>" dir="<%=dir_var%>" style="color:#FFFFFF;font-weight:600;vertical-align:middle;padding-<%=align_var%>:10px"><%=comp_t%></td>                                 
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
    <tr>
      <td width="100%" bgcolor="#060165" style="height: 18px" valign="top" >
       <table border="0" width="100%" style="height: 18px" cellspacing="0" cellpadding="0">
			<tr>
				<td width="100%" height="1" nowrap></td>
			</tr>
			<tr>
			<td bgcolor="#060165" height="18" dir="<%=dir_obj_var%>">
		   <%''/strat of news%>
		  <%sqlStr="SELECT new_ID,New_Title,New_Date FROM News"
				sqlStr=sqlStr& " WHERE New_Site_Visible=1 " 
				If trim(lang_id) = "1" Then
					sqlStr=sqlStr& " AND Category_Id=1 " 
				Else
					sqlStr=sqlStr& " AND Category_Id=2 " 
				End If
				sqlStr=sqlStr& " ORDER BY New_Date DESC,New_ID DESC"				
				set newsList=con.getRecordSet(sqlStr)
				If Not newsList.EOF Then%>
				<marquee direction="<%=align_var%>"  scrolldelay=120>
		   <%do while not newsList.EOF 
				newID=newsList("New_ID")
				newTitle=newsList("New_Title")
				newDate=newsList("New_Date")	%>
				&nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;
				<a class="homeNews" href="<%=Application("VirDir")%>netcom/news.asp?ID=<%=newID%>">>> <%=newTitle%></a>
		<%	newsList.MoveNext
				loop%>
				</marquee>	
				<%End If%>				
				<%newsList.close
				set newsList=Nothing%>
		<%'//end of news%>
		</td>
		</tr>
		<tr>
		<td width="100%" height="1" nowrap></td>
		</tr> 
		</table>
      </td>      
    </tr>      
  </table>