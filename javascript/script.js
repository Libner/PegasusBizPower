NS4 = (document.layers);
   IE4 = (document.all);
   ver4 = (NS4 || IE4);
   isMac = (navigator.appVersion.indexOf("Mac") != -1);
   isMenu = (NS4 || (IE4 && !isMac));
   function startIt(){return};
   if (!ver4) event = null;
   	var idact;
   	
function changeImage(obj,srcimg)
{
	obj.src=srcimg
}


function checkImgSize(objimg,maxwidth,maxheight)
{
	if (objimg.width > maxwidth || objimg.height > maxheight)
	{
		koeffWidth = maxwidth / objimg.width
		koeffHeight = maxheight / objimg.height
		koeff = Math.min(koeffWidth,koeffHeight)
		objimg.width = objimg.width * koeff
	}
}


//return false if email is not valid ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function checkEmail(checkThisEmail){

//  return (strEmail.indexOf(".") > 2) && (strEmail.indexOf("@") > 0);
   //return /^\w+@([\w\-]+\.)+\w{2,3}$/.test(strEmail);
var myEMailIsValid = true;
var myAtSymbolAt = checkThisEmail.indexOf('@');
var myLastDotAt = checkThisEmail.lastIndexOf('.');
var mySpaceAt = checkThisEmail.indexOf(' ');
var myLength = checkThisEmail.length;


// at least one @ must be present and not before position 2
// @yellow.com : NOT valid
// x@yellow.com : VALID

if (myAtSymbolAt < 1 ) 
 {myEMailIsValid = false}


// at least one . (dot) afer the @ is required
// x@yellow : NOT valid
// x.y@yellow : NOT valid
// x@yellow.org : VALID

if (myLastDotAt < myAtSymbolAt) 
 {myEMailIsValid = false}

// at least two characters [com, uk, fr, ...] must occur after the last . (dot)
// x.y@yellow. : NOT valid
// x.y@yellow.a : NOT valid
// x.y@yellow.ca : VALID

if (myLength - myLastDotAt <= 2) 
 {myEMailIsValid = false}


// no empty space " " is permitted (one may trim the email)
// x.y@yell ow.com : NOT valid

if (mySpaceAt != -1) 
 {myEMailIsValid = false}


return myEMailIsValid

}

/*return false if string is not number*/
function checkNumber(word)
{

   for (n=0; n < word.length; n++) 
   {
     	 if (word.charAt(n)<"0" || word.charAt(n)>"9") 
	{
          return false
      	}    
   }
return true
}



