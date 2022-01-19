<%@ Page Language="vb" AutoEventWireup="false" Codebehind="workers.aspx.vb" Inherits="bizpower_pegasus.workers"%>
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
  <meta http-equiv="content-type" content="text/html; charset=UTF-8">
  <meta name="robots" content="noindex, nofollow">
  <meta name="googlebot" content="noindex, nofollow">

  <style type="text/css">
            body { font:16px Calibri;}
        table { border-collapse:separate;  }
        td {
            margin:0;
            padding:5px;
            border:1px solid grey; 
            white-space:nowrap;
        }
        div { 
            width: auto; 
            margin-right:5em; 
            overflow-y:visible;
            padding-bottom:1px; 
        }
        .headcol {
            position:absolute; 
            width:auto; 
            right:0;
            top:auto;
        }
        .long { background:yellow; letter-spacing:1em; }
  </style>

 
  </head>
	<body>
    <form id="Form1" method="post" runat="server">
 <div dir="rtl" id="container"><table dir="rtl" id="myTable">
<asp:Repeater ID=rptWorkers Runat=server>
<HeaderTemplate>        <tr><td class="headcol">TTTT 1</td>
        <td class="headcol">TTT</td><td class="long">TTT</td><td class="long">222222222TTTT2222</td><td class="long">2222222222222</td><td class="long">2222222222222</td><td class="long">2222222222222</td><td class="long">2222222222222</td><td class="long">2222222222222</td><td class="long">33333333333333</td><td class="long">444444444444444</td></tr>
</HeaderTemplate>
<ItemTemplate>
        <tr><td class="headcol">Row 1</td>
        <td class="headcol">111111111111</td><td class="long">2222222222222</td><td class="long">2222222222222</td><td class="long">2222222222222</td><td class="long">2222222222222</td><td class="long">2222222222222</td><td class="long">2222222222222</td><td class="long">2222222222222</td><td class="long">33333333333333</td><td class="long">444444444444444</td></tr>
 </ItemTemplate>
 </asp:Repeater>
</table> 
</div>
<script src="https://ajax.googleapis.com/ajax/libs/jquery/2.2.0/jquery.min.js" type="text/javascript"></script>
 
         <script src="//code.jquery.com/jquery-1.10.2.js"></script>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
  <!-- [if lte IE 9]><script src="https://cdnjs.cloudflare.com/ajax/libs/placeholders/3.0.2/placeholders.min.js"></script><![endif] -->
  
  
<script>
//alert('ff')
 $(document).ready(function () {
     var maxWidthFirst=0
     var maxWidthSecond=0
      //alert('aa')
    $("#myTable").find("tr").find("td:first").each(function()
        {
          if($(this).width() > maxWidthFirst)
              maxWidthFirst=$(this).width()
      });
      $("#myTable").find("tr").find("td:first").next().each(function () {
          if ($(this).width() > maxWidthSecond)
              maxWidthSecond = $(this).width()
      });
    $("#myTable").find("tr").find("td:first").each(function()
    {
        $(this).width(maxWidthFirst)
       $(this).next().css('right', (maxWidthFirst + 10) + 'px').css('left', '');
        $(this).next().width(maxWidthSecond)
    })

    $("#container").css('margin-right', (maxWidthFirst + maxWidthSecond + 20) + 'px').css('left', '');
    
     $("#container").css('overflow-x','scroll');  
 })

</script>
    </form>

  </body>
</html>
