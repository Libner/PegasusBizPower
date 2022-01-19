<%@ Page Language="vb" AutoEventWireup="false" Codebehind="graphSalesControl1.aspx.vb" Inherits="bizpower_pegasus2018.graphSalesControl1"%>
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
			axisX: {
				title: "",
                labelFontSize:"14",
                fontColor: "#4F81BC",
			},
			axisY:
			{
			labelFontSize:"14"
			}
			,
			data: [ 
			  {        
        type: "column",
        name:"יעד",
        LegendText: "יעד",
       
        showInLegend: true, 
       
        dataPoints: [
       <%=label2%>
           ]
    }

,

          {        
        type: "column",
             name:"ביצוע",
             showInLegend: true,
              labelFontColor: "rgb(0,75,141)",
                 labelFontSize: "10",
          LegendText: "ביצוע",
        dataPoints: [
         <%=label1%>
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
<tr ><td  align=center><span style="font-size:25px;color:#1E9999">דוח מכירות יעד מול ביצוע שנת  <%=currentYear%></span></td></tr>


</table>

       

	<div id="chartContainer" style="height: 650px; width: 100%;">
	</div>
	</center>
</body>
</html>