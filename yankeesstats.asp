<html>

<!--This is an Active Server Page (ASP) that accesses databases and uses both a computer language VB Script and HTML.

In this case it is retrieving records from a table named MyBooks in a database called Books.asp-->

<head>

<!--<link rel="stylesheet" type="text/css" href="http://www.nystromco.com/procedures2.css">-->
<title>Yankees Stats</title>

</head>


<body>

<%

'Now we're using VB Script'

'Define variables (or fields)'

Dim MyConn, RS, ID, Player, BattingAverage, HRs, RBI, ERA

'Connect to the database'

Set MyConn=Server.CreateObject("ADODB.Connection")


MyConn.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ=D:\wwwroot\integrated-fulfillment.com\database\books.mdb"

'Execute command using SQL'


SQL = "Select * from Eliezer_Stats Order by ID"

Set RS=MyConn.Execute(SQL)

num1 = 0

'close the VB Script DB code'

%>


<!--Set up the headings for the report display-->

<table id="body1" width="1000">
  <tr>
       <td class="tablebodytop" width="10" height="20">ID </td>

	   <td class="tablebodytop" width="100" height="20">Player </td>

      <td class="tablebodytop" width="10" height="20">Batting Average</td>

      <td class="tablebodytop" width="10" height="20">HRs</td>

	  <td class="tablebodytop" width="30" height="20">RBI</td>

	  <td class="tablebodytop" width="30" height="20">ERA</td>
  </tr>
    	<tbody>

<!--Start accessing the records in the database-->

 <%While Not RS.EOF

   %>

<!--Display the records one by one-->

	   <td class="tablebodymiddle" width="10" height="16"><%=RS.Fields("ID")%>
      </td>

      <td class="tablebodymiddle" width="100" height="16"><%=RS.Fields("Player")%>
	  </td>

      <td class="tablebodymiddle" width="10" height="16"><%=RS.Fields("BattingAverage")%></td>

	   <td class="tablebodymiddle" width="10" height="16"> <%=RS.Fields("HRs")%>
	   </td>
	   <td class="tablebodymiddle" width="10" height="16"><%=RS.Fields("RBI")%>
	   </td>
	   <td class="tablebodymiddle" width="10" height="16"><%=RS.Fields("ERA")%>
	   </td>

    </tr>


<%

'Find next record'

RS.MoveNext
Wend

'End of file: close everything'

RS.Close
MyConn.Close
Set RS = Nothing
Set MyConn = Nothing
%>


	</tbody>
</table>

</body>

</html>
