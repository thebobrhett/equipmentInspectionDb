<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<!--#include file="..\Functions\HitCounter.asp"-->
<html>
<head>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>PSV Inspections Due</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="InspectionsFunctions.asp"-->
<link rel=STYLESHEET href='formstyle.css' type='text/css'>
</head>
<body>
<%
'*************
' Revision History
' 
' Keith Brooks - Tuesday, May 17, 2011
'   Creation
'*************

Dim HitCounts
Dim currentuser
Dim access
Dim printDate
Dim cn
Dim rs
Dim sqlString

'Set/get hit counts.
HitCounts = HitCounter("psv_inspections_due")

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("inspections", "psv_inspections_due", currentuser)
If access <> "none" Then
	printDate = FormatDateTime(Now,0)
%>
	<div style="text-align:center">
		<table width="100%">
			<tr>
				<td class="noprint" style="width:50%;text-align:left"><a href="#" onclick="javascript: history.go(-1);">Back</a></td>
				<td style="text-align:right;width:50%">Printed on: <%=printDate%></td>
			</tr>
			<tr>
				<td colspan="2" style="text-align:center"><h1>PSV Inspections Due</h1></td>
			</tr>
		</table>
	</div>
<%
	'Define the ado connection and recordset objects.
	set cn = CreateObject("adodb.connection")
	cn.Open = DBString
	set rs = CreateObject("adodb.recordset")

	'Get the items that have never had an inspection entered.
	sqlString = "SELECT equipment_item_tag, equipment_item_name " & _
				"FROM equipment_items i INNER JOIN psv_technical_data d " & _
				"ON i.equipment_item_id=d.equipment_item_id " & _
				"WHERE equipment_type_id=1 AND next_inspection_date is null"
	Set rs = cn.Execute(sqlString)
	'Draw the header.
	%>
	<hr />
	<table style="width:100%;border-collapse:collapse">
		<tr>
			<td colspan="3" style="text-align:left;font-size:11pt;font-weight:bold;background-color:#E7E7AB">Never Been Inspected</td>
		</tr>
	<%
	If Not rs.BOF Then
		rs.MoveFirst
		%>
		<tr>
			<th style="width:25%;border:1px solid #ABC9E7">Tagname</th>
			<th style="width:50%;border:1px solid #ABC9E7">Description</th>
			<th style="width:25%">&nbsp;</th>
		</tr>
		<%
		'Display the items.
		Do While Not rs.EOF
			%>
		<tr>
			<td style="border:1px solid #ABC9E7"><%=rs("equipment_item_tag")%></td>
			<td style="border:1px solid #ABC9E7"><%=rs("equipment_item_name")%></td>
			<td>&nbsp;</td>
		<tr>
			<%
			rs.MoveNext
		Loop
	Else
	%>
		<tr>
			<td style="width:25%">&nbsp;</td>
			<td style="width:50%;font-size:11pt">No Matching Items Found</td>
			<td style="width:25%">&nbsp;</td>
	<%
	End If
	rs.Close
	%>
	</table>
	<%
	
	'Get the items that are less than 30 days overdue.
	sqlString = "SELECT equipment_item_tag, equipment_item_name, next_inspection_date " & _
				"FROM equipment_items i INNER JOIN psv_technical_data d " & _
				"ON i.equipment_item_id=d.equipment_item_id " & _
				"WHERE equipment_type_id=1 " & _
				"AND next_inspection_date<'" & FormatMySQLDate(Date()) & "' " & _
				"AND next_inspection_date>='" & FormatMySQLDate(DateAdd("d",-30,Date())) & "'"
	Set rs = cn.Execute(sqlString)
	%>
	<hr />
	<table style="width:100%;border-collapse:collapse">
		<tr>
			<td colspan="3" style="text-align:left;font-size:11pt;font-weight:bold;background-color:#E7E7AB">< 30 Days Overdue</td>
		</tr>
	<%
	If Not rs.BOF Then
		rs.MoveFirst
		'Draw the header.
		%>
		<tr>
			<th style="width:25%;border:1px solid #ABC9E7">Tagname</th>
			<th style="width:50%;border:1px solid #ABC9E7">Description</th>
			<th style="width:25%;border:1px solid #ABC9E7">Due Date</th>
		</tr>
		<%
		'Display the items.
		Do While Not rs.EOF
			%>
		<tr>
			<td style="border:1px solid #ABC9E7"><%=rs("equipment_item_tag")%></td>
			<td style="border:1px solid #ABC9E7"><%=rs("equipment_item_name")%></td>
			<td style="border:1px solid #ABC9E7"><%=rs("next_inspection_date")%></td>
		<tr>
			<%
			rs.MoveNext
		Loop
	Else
	%>
		<tr>
			<td style="width:25%">&nbsp;</td>
			<td style="width:50%;font-size:11pt">No Matching Items Found</td>
			<td style="width:25%">&nbsp;</td>
	<%
	End If
	rs.Close
	%>
	</table>
	<%
	
	'Get the items that are between 30 and 60 days overdue.
	sqlString = "SELECT equipment_item_tag, equipment_item_name, next_inspection_date " & _
				"FROM equipment_items i INNER JOIN psv_technical_data d " & _
				"ON i.equipment_item_id=d.equipment_item_id " & _
				"WHERE equipment_type_id=1 " & _
				"AND next_inspection_date<'" & FormatMySQLDate(DateAdd("d",-30,Date())) & "' " & _
				"AND next_inspection_date>='" & FormatMySQLDate(DateAdd("d",-60,Date())) & "'"
	Set rs = cn.Execute(sqlString)
	%>
	<hr />
	<table style="width:100%;border-collapse:collapse">
		<tr>
			<td colspan="3" style="text-align:left;font-size:11pt;font-weight:bold;background-color:#E7E7AB">30 - 60 Days Overdue</td>
		</tr>
	<%
	If Not rs.BOF Then
		rs.MoveFirst
		'Draw the header.
		%>
		<tr>
			<th style="width:25%;border:1px solid #ABC9E7">Tagname</th>
			<th style="width:50%;border:1px solid #ABC9E7">Description</th>
			<th style="width:25%;border:1px solid #ABC9E7">Due Date</th>
		</tr>
		<%
		'Display the items.
		Do While Not rs.EOF
			%>
		<tr>
			<td style="border:1px solid #ABC9E7"><%=rs("equipment_item_tag")%></td>
			<td style="border:1px solid #ABC9E7"><%=rs("equipment_item_name")%></td>
			<td style="border:1px solid #ABC9E7"><%=rs("next_inspection_date")%></td>
		<tr>
			<%
			rs.MoveNext
		Loop
	Else
	%>
		<tr>
			<td style="width:25%">&nbsp;</td>
			<td style="width:50%;font-size:11pt">No Matching Items Found</td>
			<td style="width:25%">&nbsp;</td>
	<%
	End If
	rs.Close
	%>
	</table>
	<%
	
	'Get the items that are more than 60 days overdue.
	sqlString = "SELECT equipment_item_tag, equipment_item_name, next_inspection_date " & _
				"FROM equipment_items i INNER JOIN psv_technical_data d " & _
				"ON i.equipment_item_id=d.equipment_item_id " & _
				"WHERE equipment_type_id=1 " & _
				"AND next_inspection_date<'" & FormatMySQLDate(DateAdd("d",-60,Date())) & "'"
	Set rs = cn.Execute(sqlString)
	%>
	<hr />
	<table style="width:100%;border-collapse:collapse">
		<tr>
			<td colspan="3" style="text-align:left;font-size:11pt;font-weight:bold;background-color:#E7E7AB">> 60 Days Overdue</td>
		</tr>
	<%
	If Not rs.BOF Then
		rs.MoveFirst
		'Draw the header.
		%>
		<tr>
			<th style="width:25%;border:1px solid #ABC9E7">Tagname</th>
			<th style="width:50%;border:1px solid #ABC9E7">Description</th>
			<th style="width:25%;border:1px solid #ABC9E7">Due Date</th>
		</tr>
		<%
		'Display the items.
		Do While Not rs.EOF
			%>
		<tr>
			<td style="border:1px solid #ABC9E7"><%=rs("equipment_item_tag")%></td>
			<td style="border:1px solid #ABC9E7"><%=rs("equipment_item_name")%></td>
			<td style="border:1px solid #ABC9E7"><%=rs("next_inspection_date")%></td>
		<tr>
			<%
			rs.MoveNext
		Loop
	Else
	%>
		<tr>
			<td style="width:25%">&nbsp;</td>
			<td style="width:50%;font-size:11pt">No Matching Items Found</td>
			<td style="width:25%">&nbsp;</td>
	<%
	End If
	rs.Close
	%>
	</table>
	<%
	
	Set rs = Nothing
	cn.Close
	Set cn = Nothing
Else
	response.write "<h1>You don't have permission to access this page.</h1>"
	response.write "<br />"
	response.write "<a href='" & request("http_referer") & "'>Return to previous page</a>"
End If
%>
</body>
</html>
