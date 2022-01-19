<%@ Page Language="vb" AutoEventWireup="false" Codebehind="Insurance_Send.aspx.vb" Inherits="bizpower_pegasus2018.Insurance_Send" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Insurance_Send</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	<script type="text/javascript">

 $(document).ready(function(){  

  // alert("rrr")
     var uname = document.getElementById("uname").value();
     var password = document.getElementById("pwd").value();


     $('#ok').click(function(){  
         $.ajax({  
             url:'http://services.passportcard.co.il/PC_Site_DMZ_WebService/PC_Mobile_WebService.asmx',  
             type:'post',  
             dataType: 'Jsondemo',


             success: function(data) {  
                 $('#name').val(data.name);  
                 $('#email').val(data.email);  

                 var JSONObject= {
                         "uname":uname,
                         "password":password
                         };
             }  
         });  
     });  
}); 

</script>  	
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
		</form>
	</body>
</HTML>
