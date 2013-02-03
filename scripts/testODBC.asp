<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0//EN">
<HTML>

<!-- --------------------------------------------------------------------------------------
     Alentus Default Web Site

       Page.......: ODBC Connection Test Page
       Domain.....: <% response.write Request.ServerVariables("HTTP_HOST") %>
       URL........: http://<% response.write Request.ServerVariables("HTTP_HOST") %>/scripts/testODBC.asp

     Copyright 2001, Alentus Corporation
     -------------------------------------------------------------------------------------- -->

<!-- HTML Header -->
<HEAD>
  <TITLE>Welcome to <% response.write Request.ServerVariables("HTTP_HOST") %></TITLE>
  <META name='keywords' content='<% response.write Request.ServerVariables("HTTP_HOST") %>'>
</HEAD>

<STYLE>
	<!-- body	{ font-family:arial;                                    } -->
	<!-- td		{ font-family:arial; font-size:10pt; font-weight:normal } -->
	<!-- th		{ font-family:arial; font-size:10pt; font-weight:bold   } -->
	<!-- p		{ font-family:arial; font-size:10pt;                    } -->
</STYLE>

<!-- Page Body -->
<BODY bgcolor=FFFFFF>

<!-- Page Banner -->
<HR><CENTER>
<FONT size=4>Welcome to <% response.write Request.ServerVariables("HTTP_HOST") %></FONT>
<FONT size=2>
<HR>

<!-- Page Title -->
ODBC Connection Test Page.
<P>
<IMG src='/images/ODBCLogo.gif' align=center border=0>
<P>

<!-- Call the ODBC Test Routine -->
<% Call TestODBC() %>

<P>
<A href=/>Click here</A> to return to the home page.

<!-- Page Footer -->
<HR>
<P><%response.write Request.ServerVariables("HTTP_HOST") %> is hosted by:
<BR>
<A href='http://www.alentus.com'><IMG src='/images/AlentusLogo.gif' align=center border=0></A>
</FONT>
</BODY>
</HTML>

<!-- ASP Subroutines and Functions are embedded here          -->
<!--                                                          -->
<!--   They do not show up in the browser because they are    -->
<!--   filtered out at the server and executed as required    -->
<!--                                                          -->
<SCRIPT LANGUAGE=VBScript RUNAT=Server>

'---------------------------------------------------------------------------
' Subroutine   : SendLn
' Description  : Sends a line of text to the response object
'---------------------------------------------------------------------------
Sub SendLn(pMsg)
  Response.Write pMsg & chr(10)
End Sub

'---------------------------------------------------------------------------
' Subroutine   : TestODBC
' Description  : Tests the connection to the database
'---------------------------------------------------------------------------
Sub TestODBC

	dim strODBC    ' Connection string
	dim strSQL     ' SQL query string
	dim d1         ' Data connection object
	dim r1         ' Recordset for query results
	dim i          ' Counter

	' Define a connection string for the database
	'   The connect string is used to login to the database..
	strODBC = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE="&Request.ServerVariables("APPL_PHYSICAL_PATH")&"/database/artstudio.mdb"

	' Define an SQL query
	'   The query is to retrieve the ID,Name and Description fields
	'   from a table named Test, and will return the records in
	'   order of ID.
	strSQL = "SELECT ID,Name,Description FROM Test ORDER BY ID"

	' Create a database connection object
	set d1 = Server.CreateObject("ADODB.Connection")

	' Open the connection to the database
	'   The OPEN method of the database connection object is used to
	'   connect to the database as defined by the ODBC connection
	'   string.
	d1.Open strODBC

	' Execute the SQL query
	'   This will pass the SQL query to the database.  The EXECUTE
	'   method will return a recordset object containing the
	'   records that match the query.
    set r1 = d1.Execute(strSQL)

	' Check to see if any records were found
	'   If no records were returned by the query, both the BOF
	'   (Beginning of File) and EOF (End of File) properties
	'   of the recordset object will be set to TRUE.
	if (r1.EOF = True) and (r1.BOF = True) then
		SendLn "<h4>No records found</h4>"

	else

		' Set the counter variable to zero
		i=0

		' Move to the first record
		r1.movefirst

		' Start sending an HTML table to the browser
		SendLn "<table border=1 cellpadding=3>"
		SendLn "<tr><td colspan=3><b>Table Name: </b>Test</td></tr>"
		SendLn "<tr><!" + "-- Column Headings -->"
		SendLn "  <th align=right>ID</td>"
		SendLn "  <th align=left>Name</td>"
		SendLn "  <th align=left>Description</td>"

		' Loop through the record set
		'   The "while loop" will continue until the
		'   EOF property (End of File) is true.
		while r1.EOF = false

			' Send the details of the record to the browser
			'   The field values in the record are retrieved from the
			'   recordset object.  To access the field values,
			'   use the form: r1("<field name>").
			SendLn "<tr><!" + "-- Database record: " + cStr(r1("ID")) + " -->"
			SendLn  "  <td>" + cStr(r1("ID"))  + "</td>"
			SendLn  "  <td>" + r1("Name") + "</td>"
			SendLn "  <td>" + r1("Description") + "</td>"
			SendLn  "</tr>"

			' Add one to the counter variable
			i = i+1

			' Move to the next record
			'   If you forget to do this, the WHILE loop will continue
			'   forever, because EOF will never be reached
			r1.movenext
		wend

		' Send the counter total to the broswer and finish off the HTML table
		SendLn "<tr><td colspan=3><b>Total Records: </b>" + cStr(i) + "</td></tr>"
		SendLn "</table>"
	end if

End Sub

</SCRIPT>
<!-- End of ASP Script Subroutines and Functions              -->
