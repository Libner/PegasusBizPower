<%@ Language=VBScript %>
<!-- METADATA TYPE="TypeLib" UUID="{50D94450-589A-409B-819A-4F88B151C134}"-->
<% 
  Dim AuthenticationStr
 AuthenticationStr = session("AuthenticationStr")

  response.ContentType = "image/png"
  textToGIF(AuthenticationStr)
  
Function textToGIF(inText)
	
	on error resume next

	Dim objImageGen 
	Set objImageGen = Server.CreateObject("softartisans.ImageGen") ' As New ImageGen
  
    objImageGen.CreateGradientImage 90, 32, rgb(204,204,204), rgb(247,247,247), 0
	objImageGen.ColorResolution = saiTrueColors
	
	randomize
  
  	Dim Fonts
	Fonts = array("Times New Roman","Arial","Book Antiqua","Comic Sans MS","Haettenschweiler","Monotype Corsiva","Impact")

	 'set random fonts.
	Set objFont = objImageGen.Font	
    objFont.Name =  "Lucida Console"
	objFont.Weight = 800	
	objFont.height = 16
	objFont.orientation = -8
	objFont.Color = rgb(0,0,0)
	if rnd>0.5 then objFont.Weight = 800
	
	objImageGen.AntiAliasFactor = saAntiAliasBest
	objImageGen.Text = inText
    objImageGen.DrawTextOnImage 7, 8, 83, 0
 
	objImageGen.SaveImage saiBrowser, saiPNG, "captcha.png"
	Set objImageGen = Nothing	
End Function %>