<meta http-equiv="Content-Language" content="en-us">
<%@Language=VBScript%>
<%Response.Buffer=True%>

<html>
<body>
	
<!--#include file="adovbs.inc"-->
	
<%

'*** This script adds a new record to the database Books to the Blog table ***

'*** Define the variables/fields ***

Dim MyConn, RS, ID, Date1, BlogTitle, Blog, Category

'*** Equate the fields received from the form to variables ***

strDate1 = Request.Form("Date1")
strBlogTitle = Request.Form("BlogTitle")
strBlog = Request.Form("Blog")
strCategory = Request.Form("Category")



Set MyConn=Server.CreateObject("ADODB.Connection")

Set RS=Server.CreateObject("ADODB.RecordSet")
RS.CursorType = adOpenForwardOnly
RS.LockType = adLockPessimistic

'*** Select Connection Method ***
MyConn.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ=D:\wwwroot\integrated-fulfillment.com\database\books.mdb"


RS.Open "Select * from Blog", MyConn, , , adCMDText

RS.AddNew
RS("Date1") = strDate1
RS("BlogTitle") = strBlogTitle
RS("Blog") = strBlog
RS("Category") = strCategory

RS.Update


RS.Close
MyConn.Close
Set RS = Nothing
Set MyConn = Nothing

Response.Redirect "add_record.html"
%>
<p>&nbsp;</p>

</body>
</html>