<html>
<!--#include file="../connect.asp"-->
<!--#include file="../reverse.asp"-->
<!--#include file="checkWorker.asp"-->
<%set con = nothing%>
<head>
<title>BizPower</title>
<meta charset="windows-1255">
<link href="../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<body marginwidth="0" marginheight="0" topmargin="0" leftmargin="0" rightmargin="0" bgcolor="#E5E5E5">
<div align="center"><center>
<table border="0" width="100%" bgcolor="#FFFFFF" height="70" cellspacing="0" cellpadding="0" ID="Table1">
  <tr>
    <td><img src="../../images/leumi_top.jpg"></td>
  </tr>
</table>
</center></div><div align="center"><center>

<div align="center"><center>
<table border="0" width="100%" cellspacing="0" cellpadding="0" bgcolor="#060165">
  <%numOftab = 6%>
  <tr><td width=100% align="right"><!--#include file="../top.asp"--></td></tr>
</table>
</center></div>
<table border="0" width="100%" bgcolor="#FFFFFF" height="300" cellspacing="0" cellpadding="0" background="../images/bgr_list.jpg">
   <tr>
    <td width="118" valign="top"><img src="../../images/pict1_netcom.jpg" width="118" height="301"></td>
    <td valign="top" width="202"><img src="../../images/pict2_netcom.jpg" height="301"></td>
    <td width="100%" valign="top" align="right">
     <table border="0" width="100%" cellspacing="0" cellpadding="0">
<!-- MAIN -->    
      <tr><td width="100%" height="15"></td></tr>
      <tr><td width="100%">
        <table width=93% align="center" border="0" cellspacing="0" cellpadding="2">
           <tr>
              <td align="right" width="100%">.<font style="color:#003497;"><b><u>1000</u></b></font>&nbsp;-&nbsp;<b>���� ������ ������</b></td>      
           </tr>
           <tr><td width="100%" height="5"></td></tr>
           <tr>
               <td align="right" bgcolor="#EBEBEB" dir="rtl"><font style="color:#000000;font-size:11pt;"><b>������ ����� ������ ���� �����</b>(���� ����).</font>&nbsp;</td>      
           </tr>
           <tr><td width="100%" height="5"></td></tr>
			<tr>
			<td width="100%" valign="top" align="right">
			<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#7d7d7c;">��� 1. </span><a href="pages/default.asp"><span style="font-size:11pt;font-weight:bold;color:#003497;text-decoration:none">���� �����</span></a></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#000000;">��� ���� �����</span>, ����� ����� ��
���� ���� ������ ���� ���� ���� ����� ���� �����, ������ ������� ����.</td></tr>
           <tr><td width="100%" height="3"></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#7d7d7c;">��� 2. </span><a href="products/questions.asp"><span style="font-size:11pt;font-weight:bold;color:#003497;text-decoration:none">���� ����</span></a></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#000000;">��� ���� ����</span> ��� ����� �����
��������� (����� �����, ���� ����, ��� ����&quot;�).</td></tr>
           <tr><td width="100%" height="3"></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#7d7d7c;">��� 3. </span><a href="companies/choose.asp"><span style="font-size:11pt;font-weight:bold;color:#003497;text-decoration:none">������ ������</span></a></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#000000;">������ �� ������</span> ����
������ ������ ���&quot;� ������ ����� ���� �� ������ ����� <span dir=ltr>EXCEL</span></td></tr>
           <tr><td width="100%" height="3"></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#7d7d7c;">��� 4. </span><a href="products/products.asp"><span style="font-size:11pt;font-weight:bold;color:#003497;text-decoration:none">���� �������</span></a></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#000000;">���� ����</span> ���� ����� �� �� ���
���� ���� ������� �������.</td></tr>
           <tr><td width="100%" height="3"></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#7d7d7c;">��� 5. </span><a href="admin_DB/default.asp"><span style="font-size:11pt;font-weight:bold;color:#003497;text-decoration:none">������</span></a></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#000000;">���� �������</span> (������) �� ����� �����
������ ������.</td></tr>
           <tr><td width="100%" height="3"></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#7d7d7c;">��� 6. </span><a href="reports/reports.asp"><span style="font-size:11pt;font-weight:bold;color:#003497;text-decoration:none">�����</span></a></td></tr>
			<tr><td align="right" dir="rtl"><span style="font-size:10pt;font-weight:bold;color:#000000;">����� ������� �������</span> �� �����
�����, ����� ��������� ������ ����.</td></tr>
           <tr><td width="100%" height="5"></td></tr>
           <tr><td width="100%" align="right" dir="rtl">
���� ����� ������� ������ ���
������ �- <span dir=ltr>Netcommunicator</span></td></tr>
           <tr><td width="100%" align="right" dir="rtl">
<b>���� ����� �������</b> ���� �������
����� ������ ���' <span dir=ltr><b>04-8770282</b></span>. ��� ������, ��� ����� ������, ���� ������� ������
��� ��� ����� ������ &#8211; ���� �� ���� �������� �� ������ �� ������.
           </td></tr>
           <tr><td width="100%" height="10"></td></tr>
			</table>
			</td>
			</tr>  
        </table>
       </td>
      </tr>  
<!-- END MAIN -->        
     </table>
    </td>
    <td width="17" valign="top" align="right"></td>
  </tr>
</table>
</center></div>
</body>
</html>
