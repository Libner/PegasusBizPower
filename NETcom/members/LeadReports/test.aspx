<%@ Page Language="vb" AutoEventWireup="false" Codebehind="test.aspx.vb" Inherits="bizpower_pegasus.test"%>
<!DOCTYPE html>
<html>
<head>
<script type="text/javascript">
window.onload = function () {
      var chart = new CanvasJS.Chart("chartContainer");
    chart.options.axisY = {};
    chart.options.title = { 
     fontSize: "25",
    text: "<%=title%>" };

    var series1 = { //dataSeries - first quarter
        type: "column",
        color:"#1E9EC1",
        name: "כמות טופס צור קשר מקוצר וצור קשר רגיל",
        showInLegend: true
    };

    var series2 = { //dataSeries - second quarter
        type: "column",
        color:"#3B579D",
        name: "כמות טפסי המתעניין שנפתחו",
        showInLegend: true
    };
 var series3 = { //dataSeries - second quarter
        type: "column",
        color:"#8500FF",
        name: "כמות טפסי המתעניין בסטטוס סגור נרשם",
        showInLegend: true
    };

    chart.options.data = [];
    chart.options.data.push(series1);
    chart.options.data.push(series2);
    chart.options.data.push(series3);

    series1.dataPoints = [
        <%=label1%>
         
    ];

    series2.dataPoints = [
     <%=label2%>
    
    ];
 series3.dataPoints = [
  <%=label3%>
          
    ];
    chart.render();
}
</script>
<script type="text/javascript" src="https://canvasjs.com/assets/script/canvasjs.min.js"></script>
</head>

<body>
    <div id="chartContainer" style="height: 800px; width: 100%;">
    </div>
</body>
</html>