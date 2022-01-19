<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
Function CreateGUID()
  Randomize Timer
  Dim tmpCounter,tmpGUID
  Const strValid = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  For tmpCounter = 1 To 20
    tmpGUID = tmpGUID & Mid(strValid, Int(Rnd(1) * Len(strValid)) + 1, 1)
  Next
  CreateGUID = tmpGUID
End Function%>
<%'response.write CreateGUID()
'			response.end
set rs_supplier = con.GetRecordSet("select supplier_Id from Suppliers")
    if not rs_supplier.eof then 
		do while not rs_supplier.eof
		supplierId=rs_supplier("supplier_Id")
    GUID=CreateGUID()
sqlUpd="  update Suppliers set GUID='"& GUID &"' where supplier_Id=" &supplierId
    		con.executeQuery(sqlUpd)
    rs_supplier.MoveNext
		loop
end if

%>
