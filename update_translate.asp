<!--#include file="netcom/connect.asp" -->
<!--#include file="netcom/reverse.asp" -->
<%

con.executeQuery("DELETE FROM translate")

FileName = "translate.xls"
Const adOpenStatic = 3
Const adLockPessimistic = 2

Set cnnExcel = Server.CreateObject("ADODB.Connection")
cnnExcel.Open "DBQ=" & Server.MapPath("download/"& FileName ) & ";" & _
	"DRIVER={Microsoft Excel Driver (*.xls)};"

' Same as any other data source.
' FYI: mytable is my named range in the Excel file
Set rstExcel = Server.CreateObject("ADODB.Recordset")
rstExcel.Open "SELECT page_id, lang_id, word_id, word FROM translate;", cnnExcel, adOpenStatic, adLockPessimistic
' Get a count of the fields and subtract one since we start
' counting from 0.
iCols = rstExcel.RecordCount
'Response.Write "iCols=" & iCols & "<br>"
if not rstExcel.EOF then
	rstExcel.MoveNext ' this is row of selolar number and name on hebrew, don't add.
end if

'//loop over the excell record
'//insert just the correct record which have valid cellular and nort null name
Do while not rstExcel.EOF 
	page_id = trim(rstExcel("page_id"))
	lang_id = trim(rstExcel("lang_id"))
	word_id = trim(rstExcel("word_id"))
	word = trim(rstExcel("word"))	
	
	If IsNumeric(page_id) And IsNumeric(lang_id) And IsNumeric(word_id) Then
	sqlStr = "Insert Into translate (page_id, lang_id, word_id, word) Values (" &_
	page_id & "," & lang_id & "," & word_id & ",'" & sFix(word) & "')"
	con.ExecuteQuery(sqlStr)
	End If
	
	rstExcel.MoveNext
Loop


'end of reading records from excell file			
 
rstExcel.Close
Set rstExcel = Nothing

cnnExcel.Close
Set cnnExcel = Nothing
%>