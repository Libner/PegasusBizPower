<%@ Page Language="vb" AutoEventWireup="false" Codebehind="testPopup.aspx.vb" Inherits="bizpower_pegasus.testPopup"%>
<!DOCTYPE html>
<html>  <head>
    <title>testPopup</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
   <script src="../../../javaScript/jquery-1.11.2.min.js"></script>	
    <script type="text/javascript" src="../../../javaScript/jquery.bpopup.min.js?v=<%=Application("sversion")%>"></script>			
	<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script>

   

$(document).on("click", ".popup-iframe", function( event ) {
	var href = $(this).attr("popup-url");
    if (typeof (href) == "undefined") {
		href = "";
    }		
    
	var iframeAttr = $(this).attr("data-iframe");
    if (typeof (iframeAttr) == "undefined") {
		iframeAttr = "";
    }	
        
	//alert(href);	
	var param_obj = {href: href, iframeAttr: iframeAttr};
	alert(param_obj)
	parent.popupOpen(param_obj);
        
	event.preventDefault();
        
	return false;
});

/*function popupOpen(param_obj){
	
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
			modalClose: true,
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
	
	//alert('closeBPopup Finish');	
}*/
</script>

  </head>
  <body MS_POSITIONING="GridLayout">

    <form id="Form1" method="post" runat="server">
    <a class="popup-iframe" href="" popup-url="http://bluto/bizpower_pegasus/netcom/members/messages/ViewMessage.asp?messageId=147" data-iframe='scrolling="yes" frameborder="0" width="800" height="500"'>
										<span>click</span>
				</a>
<div id="customPopup1">
        <span class="button b-close"><span><i class="fa fa-times">X</i></span></span>
        <div class="content1">
        </div>
</div>
		
    </form>

  </body>
</html>
