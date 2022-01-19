<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ReportGraph.aspx.vb" Inherits="bizpower_pegasus.ReportGraph"%>
<!DOCTYPE HTML>
<html>

<head>  
<style>
BODY
{
    PADDING-RIGHT: 0px;
    PADDING-LEFT: 0px;
    PADDING-BOTTOM: 0px;
    MARGIN: 0pt;
    FONT-FAMILY: Arial;
    COLOR: #000000;
    PADDING-TOP: 0px;
    BACKGROUND-COLOR: #ffffff
    
}
TD
{
    FONT-WEIGHT: normal;
    FONT-SIZE: 12pt;
    FONT-FAMILY: Arial;
    COLOR: #000000;
}
</style>
	<script type="text/javascript">
	window.onload = function () {
		var chart = new CanvasJS.Chart("chartContainer",
		{
			theme: "theme3",
                        animationEnabled: true,
			title:{
				 text: "<%=title%>", 
		  fontWeight: "bolder",
        fontColor: "#008B8B",
        fontfamily: "arial",        
        fontSize: 25,
        padding: 10        
      },
			toolTip: {
				shared: true
			},			
			axisY: {
				title: "ציון %"
			},
			axisX:
			{
			labelFontSize:"12"
			}
			,
			data: [ 
			  {        
        type: "column",
        name:"ציון טיול",
        LegendText: "ציון טיול",
       
        showInLegend: true, 
       
        dataPoints: [
       <%=label1%>
           ]
    }

,

          {        
        type: "column",
             name:"ציון מדריך בטיול",
             showInLegend: true,
              labelFontColor: "rgb(0,75,141)",
                 labelFontSize: "10",
          LegendText: "ציון מדריך בטיול",
        dataPoints: [
         <%=label2%>
      ]
    }
			
		
			
			],
          legend:{
            cursor:"pointer",
            itemclick: function(e){
              if (typeof(e.dataSeries.visible) === "undefined" || e.dataSeries.visible) {
              	e.dataSeries.visible = false;
              }
              else {
                e.dataSeries.visible = true;
              }
            	chart.render();
            }
          },
        });

chart.render();
}
</script>
  <script type="text/javascript" src="https://canvasjs.com/assets/script/canvasjs.min.js"></script>
</head>
<body>
<center>
<table cellpadding=0 cellspacing=0 border=0 width =100%>
<tr><TD height=20></td></tr>
<tr ><td  align=center>נכון לתאריך: <%=Now()%> </td></tr>

</table>

       

	<div id="chartContainer" style="height: 800px; width: 100%;">
	</div>
	</center>
</body>
</html>