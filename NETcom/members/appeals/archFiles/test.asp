<!DOCTYPE html>
<html>
<body>

<%
res=0
Dim myarray()		
Redim myarray(3) 	
txt="This is a beautiful day!אין"


myarray(0) = "אין מענה"
myarray(1) = "אין תשובה"
myarray(2) = "ממתינה"
myarray(3) = "לא עונה"
response.Write Ubound(myarray)
'response.end
For ii=0 To Ubound(myarray)
response.WRITE(res &":" &InStr(txt,myarray(ii)) &":"& myarray(ii) &"<br>")
if InStr(txt,myarray(ii))>0 then
	res=1
  Exit For
  end if
next

response.write("res="&res)

%>  

</body>
</html>