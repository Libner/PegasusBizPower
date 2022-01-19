<div id="customPopup1">
        <span class="button b-close"><span><i class="fa fa-times"></i></span></span>
        <div class="content1">
        </div>
</div>


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
	popupOpen(param_obj);
        
	event.preventDefault();
        
	return false;
});

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
}
</script>

<style>
/*start of popup jquery*/
#customPopup1,#customPopup2,#customPopup3,.bMulti {
    background-color: #fff;
    border-radius: 5px 5px 5px 5px;
    box-shadow: 0 0 25px 5px rgba(0, 0, 0, 0.5);
    color: #111;
    display: none;
    min-width: 450px;
    padding: 25px;
	max-width: 900px;
	direction: rtl;    
}

#popup,.bMulti {
    min-height: 250px
}

#customPopup1 iframe, #customPopup2 iframe, #customPopup3 iframe {
    background: url('/images/popupLoader.gif') center center no-repeat;
    min-height: 240px;
    min-width: 450px
}

.loading {
    background: url('/images/popupLoader.gif') center center no-repeat
}

.button {
    background-color: #990000;
    border-radius: 10px;
    box-shadow: 0 2px 3px rgba(0,0,0,0.3);
    color: #fff;
    cursor: pointer;
    display: inline-block;
    padding: 10px 20px;
    text-align: center;
    text-decoration: none
}

.button.small {
    border-radius: 15px;
    float: right;
    margin: 22px 5px 0;
    padding: 6px 15px
}

.button:hover {
    background-color: #1e1e1e
}

.button>span {
    font-size: 140%
}

.button.b-close,.button.bClose {
    border-radius: 5px 5px 5px 5px;
    box-shadow: none;
    font: bold 131% sans-serif;
    padding: 1px 10px 2px;
    position: absolute;
    right: -15px;
    top: -15px
}
/*end of popup jquery*/	
</style>