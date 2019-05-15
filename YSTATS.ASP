<meta http-equiv="Content-Language" content="en-us">
<%@Language=VBScript%>
<%Response.Buffer=True%>

<html>
<body>
	
<!--#include file="adovbs.inc"-->
	
<%

'*** This script adds a new record to the database Books to the Blog table ***

'*** Define the variables/fields ***

Dim MyConn, RS, ID, Player, BattingAverage, HRs, RBI, ERA  

'*** Equate the fields received from the form to variables ***

strPlayer = Request.Form("Player")
strBattingAverage = Request.Form("BattingAverage")
strHRs = Request.Form("HRs")
strRBI = Request.Form("RBI")
strERA = Request.Form("ERA")



Set MyConn=Server.CreateObject("ADODB.Connection")

Set RS=Server.CreateObject("ADODB.RecordSet")
RS.CursorType = adOpenForwardOnly
RS.LockType = adLockPessimistic

'*** Select Connection Method ***
MyConn.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ=D:\wwwroot\integrated-fulfillment.com\database\books.mdb"


RS.Open "Select * from Eliezer_Stats", MyConn, , , adCMDText

RS.AddNew
RS("Player") = strPlayer
RS("BattingAverage")= strBattingAverage
RS("HRs") = strHRs
RS("RBI") = strRBI
RS("ERA") = strERA
RS.Update


RS.Close
MyConn.Close
Set RS = Nothing
Set MyConn = Nothing

Response.Redirect "New_Entry.html"
%>
<p>&nbsp;</p>

</body>
</html>