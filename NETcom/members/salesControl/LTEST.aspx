<%@ Page Language="vb" AutoEventWireup="false" Codebehind="LTEST.aspx.vb" Inherits="bizpower_pegasus2018.LTEST"%>

<!DOCTYPE html>
<html>
<head>
<style class="cp-pen-styles">.twoToneCenter {
  text-align: center;
  margin: 1em 0;
}
.twoToneButton {
  display: inline-block;
  outline: none;
  padding: 10px 20px;
  line-height: 1.4;
  background: #212121;
  background: linear-gradient(to bottom, #545454 0%, #474747 50%, #141414 51%, #1b1b1b 100%);
  border-radius: 4px;
  border: 1px solid #000000;
  color: #dadada;
  text-shadow: #000000 -1px -1px 0px;
  position: relative;
  transition: padding-right 0.3s ease;
  font-weight: 700;
  box-shadow: 0 1px 0 #6e6e6e inset, 0px 1px 0 #3b3b3b;
}
.twoToneButton:hover {
  box-shadow: 0 0 10px #080808 inset, 0px 1px 0 #3b3b3b;
  color: #f3f3f3;
}
.twoToneButton:active {
  box-shadow: 0 0 10px #080808 inset, 0px 1px 0 #3b3b3b;
  color: #ffffff;
  background: #080808;
  background: linear-gradient(to bottom, #3b3b3b 0%, #2e2e2e 50%, #141414 51%, #080808 100%);
}
.twoToneButton.spinning {
  background-color: #212121;
  padding-right: 40px;
}
.twoToneButton.spinning:after {
  content: '';
  right: 6px;
  top: 50%;
  width: 0;
  height: 0;
  box-shadow: 0px 0px 0 1px #080808;
  position: absolute;
  border-radius: 50%;
  -webkit-animation: rotate360 0.5s infinite linear, exist 0.1s forwards ease;
          animation: rotate360 0.5s infinite linear, exist 0.1s forwards ease;
}
.twoToneButton.spinning:before {
  content: "";
  width: 0px;
  height: 0px;
  border-radius: 50%;
  right: 6px;
  top: 50%;
  position: absolute;
  border: 2px solid #000000;
  border-right: 3px solid #27ae60;
  -webkit-animation: rotate360 0.5s infinite linear, exist 0.1s forwards ease;
          animation: rotate360 0.5s infinite linear, exist 0.1s forwards ease;
}
@-webkit-keyframes rotate360 {
  100% {
    -webkit-transform: rotate(360deg);
            transform: rotate(360deg);
  }
}
@keyframes rotate360 {
  100% {
    -webkit-transform: rotate(360deg);
            transform: rotate(360deg);
  }
}
@-webkit-keyframes exist {
  100% {
    width: 15px;
    height: 15px;
    margin: -8px 5px 0 0;
  }
}
@keyframes exist {
  100% {
    width: 15px;
    height: 15px;
    margin: -8px 5px 0 0;
  }
}
body {
  background: #212121;
}
</style></head>
<body>
 
<div class="twoToneCenter">
    
    <button class="twoToneButton">Sign In</button>
    
</div>
<script src='https://static.codepen.io/assets/common/stopExecutionOnTimeout-de7e2ef6bfefd24b79a3f68b414b87b8db5b08439cac3f1012092b2290c719cd.js'></script><script src='//cdnjs.cloudflare.com/ajax/libs/jquery/2.1.3/jquery.min.js'></script>
<script >$(function () {

    var twoToneButton = document.querySelector('.twoToneButton');

    twoToneButton.addEventListener("click", function () {
        twoToneButton.innerHTML = "Signing In";
        twoToneButton.classList.add('spinning');

        setTimeout(
        function () {
            twoToneButton.classList.remove('spinning');
            twoToneButton.innerHTML = "Sign In";

        }, 6000);
    }, false);

});
</script>
</body></html>