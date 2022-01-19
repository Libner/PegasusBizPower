<%@ Page Language="vb"  AutoEventWireup="false" Codebehind="reportYearGraph_P2932.aspx.vb" Inherits="bizpower_pegasus2018.reportYearGraph"%>
<HTML>
	<HEAD>
		<DOCTYPE HTML>
		<script type="text/javascript">
  window.onload = function () {
    var chart = new CanvasJS.Chart("chartContainer",
    {
       title:{
        text: "<%=title%>", 
        	
        fontWeight: "bolder",
        fontColor: "#008B8B",
        fontfamily: "tahoma",        
        fontSize: 25,
        padding: 10        
      },
     	axisX: {
					labelFontSize: 14,
		            indexLabelFontSize:12
			},
	axisY: {
					labelFontSize: 14,
		            indexLabelFontSize:12
			},
      data: [
   
   {        
        type: "column",
        name:"<%=Year(DateAdd("yyyy",-2,dateStart))%>",
        LegendText: "<%=Year(DateAdd("yyyy",-2,dateStart))%>",
        showInLegend: true, 
       
        dataPoints: [
       <%=label2%>
           ]
    }

,

          {        
        type: "column",
             name:"<%=Year(DateAdd("yyyy",-1,dateStart))%>",
             showInLegend: true, 
          LegendText: "<%=Year(DateAdd("yyyy",-1,dateStart))%>",
        dataPoints: [
         <%=label1%>
      ]
    }
,
          {        
        type: "column",
             name:"<%=Year(dateStart)%>",
           showInLegend: true,   
              LegendText: "<%=Year(dateStart)%>",
        dataPoints: [
       <%=label%>
      ]
    }

      ]
    });

    chart.render();
  }
		</script>
		<script type="text/javascript" src="../../../graphs/canvasjs.min.js"></script>
	</HEAD>
	<body>
		<div id="chartContainer" style="HEIGHT: 700px; WIDTH: 100%">
		</div>
	</body>
</HTML>
