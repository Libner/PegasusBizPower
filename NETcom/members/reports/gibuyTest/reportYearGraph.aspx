<%@ Page Language="vb"  AutoEventWireup="false" Codebehind="reportYearGraph.aspx.vb" Inherits="bizpower_pegasus.reportYearGraph"%>
<DOCTYPE HTML>
<html>
<head>  
  <script type="text/javascript">
  window.onload = function () {
    var chart = new CanvasJS.Chart("chartContainer",
    {
       title:{
        text: "דוח השווה", 
        fontWeight: "bolder",
        fontColor: "#008B8B",
        fontfamily: "tahoma",        
        fontSize: 25,
        padding: 10        
      },

      data: [
   
   {        
        type: "column",
        name:"2015",
        LegendText: "2015",
        showInLegend: true, 
       
        dataPoints: [
         {label: "אוסטריה ושוויץ (האלפים)", y: 41,indexLabel: "41/100"},
         {label: "איטליה", y: 343,indexLabel: "343/100"},
         {label: "אלבניה", y: 46,indexLabel: "46/100"}
           ]
    }

,

          {        
        type: "column",
             name:"2016",
             showInLegend: true, 
          LegendText: "2016",
        dataPoints: [
         {label: "אוסטריה ושוויץ (האלפים)", y: 45,indexLabel: "45/100" },
         {label: "איטליה", y: 296,indexLabel: "296/100"},
         {label: "אלבניה", y: 68,indexLabel: "68/100"}
      ]
    }
,
          {        
        type: "column",
             name:"2017",
           showInLegend: true,   
              LegendText: "2017",
        dataPoints: [
         {label: "אוסטריה ושוויץ (האלפים)", y: 21,indexLabel: "21/100" },
         {label: "איטליה", y: 241,indexLabel: "241/100"},
         {label: "אלבניה", y: 33,indexLabel: "33/100"}
      ]
    }

      ]
    });

    chart.render();
  }
  </script>
<script type="text/javascript" src="../../../graphs/canvasjs.min.js"></script></head>
<body>
  <div id="chartContainer" style="height: 300px; width: 100%;">
  </div>
</body>
</html>
