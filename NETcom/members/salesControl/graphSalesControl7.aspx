<%@ Page Language="vb" AutoEventWireup="false" Codebehind="graphSalesControl7.aspx.vb" Inherits="bizpower_pegasus2018.graphSalesControl7"%>
<!DOCTYPE HTML>
<HTML>
	<HEAD>
		<style> BODY { PADDING-RIGHT: 0px; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; MARGIN: 0pt; FONT-FAMILY: Arial; COLOR: #000000; PADDING-TOP: 0px; BACKGROUND-COLOR: #ffffff }
	TD { FONT-WEIGHT: normal; FONT-SIZE: 12pt; FONT-FAMILY: Arial; COLOR: #000000; }
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
                labelFontSize:"12"
			},
			axisY:
			{
			labelFontSize:"12"
			}
			,
			data: [ 
			  {        
        type: "column",
        name:" קישוריות מכירה "+ <%=currentYear-1%>,
        LegendText:" קישוריות מכירה"+ <%=currentYear-1%>,
       
        showInLegend: true, 
       
        dataPoints: [
       <%=label2%>
           ]
    }

,

          {        
        type: "column",
             name:" קישוריות מכירה "+ <%=currentYear%>,
             showInLegend: true,
              labelFontColor: "rgb(0,75,141)",
                 labelFontSize: "10",
          LegendText: " קישוריות מכירה"+ <%=currentYear%>,
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
	</HEAD>
	<body>
		<center>
			<table cellpadding="0" cellspacing="0" border="0" width="100%">
				<tr>
					<TD height="20"></TD>
				</tr>
				<tr>
					<td align="center">נכון לתאריך:
						<%=Now()%>
					</td>
				</tr>
				<tr>
					<td align="center"><span style="FONT-SIZE:25px;COLOR:#1e9999">קישוריות מכירה <%=currentYear-1%> לעומת <%=currentYear%> </span></td>
				</tr>
			</table>
			<div id="chartContainer" style="HEIGHT: 650px; WIDTH: 100%">
			</div>
		</center>
	</body>
</HTML>
