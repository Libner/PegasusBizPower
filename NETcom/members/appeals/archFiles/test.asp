<!DOCTYPE html>
<html>
<body>

<%
res=0
Dim myarray()		
Redim myarray(3) 	
txt="This is a beautiful day!���"


myarray(0) = "��� ����"
myarray(1) = "��� �����"
myarray(2) = "������"
myarray(3) = "�� ����"
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