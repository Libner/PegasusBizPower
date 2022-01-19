<%Response.Buffer = False%>
<html>
<head>
<title>Bizpower Administration</title>
<meta charset="windows-1255">
<link href="../admin.css" rel="STYLESHEET" type="text/css">
</head>
<body>
<div align="center">
<table border="1" style="border-collapse:collapse;" width="350" bordercolor="#000000" cellspacing="0" cellpadding="0">
  <tr height=25>
    <td bgcolor="#D2D2D2" align="center" class="12normalB" valign="middle" nowrap>כניסה לניהול אתר</td>
  </tr>
<tr><td bgcolor="#F4F4F4">
<form name=form1 id=form1 method="post" action="choose.asp">
  <table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td align="right" height="15" colspan="2"></td>
    </tr>
    <tr>
      <td width="50%" align="right"><div align="right"><p><input class="texts" type="text" name="username" id="username" size="10" tabindex="1"><font face="Arial (hebrew)" size="2"></font></td>
      <th width="50%" class="10normal" align="left">&nbsp;: שם עובד</th>
    </tr>
    <tr align="center">
      <td align="center" height="15"><div align="right"><p><input class="texts" type="password" name="password" id="password" size="10" tabindex="2"></td>
      <th class="10normal" align="left">&nbsp;: סיסמא</th>
    </tr>
    <tr>
      <td align="center" colspan="2" height="15"></td>
    </tr>
    <tr align="center">
      <td colspan=2 align="center"><input class="but" type="submit" value="כניסה" name="submit1" id="submit1" class="but_browse"></td>
    </tr>
    <tr align="center">
      <td width="50%" align="center" colspan="2"></td>
    </tr>
  </table>
</form>
</td></tr></table>
</div>
</body>
<%'end if%>
</html>
<html></html>