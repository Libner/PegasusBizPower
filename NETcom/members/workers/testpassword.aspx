<%@ Page Language="vb" AutoEventWireup="false" Codebehind="testpassword.aspx.vb" Inherits="bizpower_pegasus2018.testpassword"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
    <title>testpassword</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
       <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
   <script>function createstars(n) {
  return new Array(n+1).join("*")
}


$(document).ready(function() {

  var timer = "";

 // $(".panel").append($('<input type="text" class="hidpassw" />'));

  $(".hidpassw").attr("name", $(".pass").attr("name"));

  $(".pass").attr("type", "text").removeAttr("name");

  $("body").on("keypress", ".pass", function(e) {
    var code = e.which;
    if (code >= 32 && code <= 127) {
      var character = String.fromCharCode(code);
      $(".hidpassw").val($(".hidpassw").val() + character);
    }


  });

  $("body").on("keyup", ".pass", function(e) {
    var code = e.which;

    if (code == 8) {
      var length = $(".pass").val().length;
      $(".hidpassw").val($(".hidpassw").val().substring(0, length));
    } else if (code == 37) {

    } else {
      var current_val = $('.pass').val().length;
      $(".pass").val(createstars(current_val - 1) + $(".pass").val().substring(current_val - 1));
    }

    clearTimeout(timer);
    timer = setTimeout(function() {
      $(".pass").val(createstars($(".pass").val().length));
    }, 200);

  });

});
</script>
 
      </head>
  <body MS_POSITIONING="GridLayout">

    <form id="Form1" method="post" runat="server">

<div class="panel">
  <input type="password" name="paswd" class="pass" />
</div>

<input type="hidden" name="pass" class="hidpassw"/>
    </form>
  </body>
</html>
