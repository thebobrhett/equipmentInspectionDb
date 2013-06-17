<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<meta http-equiv="X-UA-Compatible" content="IE=8" />
<html>
<head>
<script language="javascript">
</script>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>PMs</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="InspectionsFunctions.asp"-->
<link rel=STYLESHEET href='formstyle.css' type='text/css'>
<style>
table {page-break-inside:auto}
tr {page-break-inside:avoid;
	page-break-after:auto}
td {border:1px solid black;
	text-align:left;
	padding:3px;
	font-size:7pt}
th {vertical-align:bottom;
	font-size:7pt}
</style>
</head>
<%
'*************
' Revision History
' 
' Keith Brooks - Thursday, December 29, 2011
'   Creation.
'*************

Dim sqlString
Dim cn
Dim rs
Dim rs2
Dim currentuser
Dim access
Dim itemID
Dim rowCount
Dim count
Dim equipment_item_tag
Dim equipment_item_description
Dim pm_priority
Dim equipment_items_subitem_description
Dim test_front_bearing
Dim test_rear_bearing
Dim test_top

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("inspections", "pm_manual_inspection", currentuser)
If access <> "none" Then

'	Response.Write "<body style='background-color:white' onload='window.print();window.close();'>"
	Response.Write "<body style='background-color:white'>"
		
	'Define the ado connection and recordset objects.
	set cn = CreateObject("adodb.connection")
	cn.CursorLocation = 3
	cn.Open = DBString
	set rs = CreateObject("adodb.recordset")
	Set rs2 = CreateObject("adodb.recordset")

	Response.Write "<form id='form1' name='form1' action='inspectionaction.asp' method='post'>"
	
	'Draw header.
	Response.Write "<table style='width:100%;border-collapse:collapse'>"
	Response.Write "<col width='5%' />"
	Response.Write "<col width='20%' />"
	Response.Write "<col width='15%' />"
	Response.Write "<col width='5%' />"
	Response.Write "<col width='5%' />"
	Response.Write "<col width='5%' />"
	Response.Write "<col width='5%' />"
	Response.Write "<col width='20%' />"
	Response.Write "<col width='5%' />"
	Response.Write "<col width='5%' />"
	Response.Write "<col width='10%' />"
	Response.Write "<thead>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' class='noprint' colspan='11' style='font-size:10pt;text-align:center'>"
	Response.Write "<a href='javascript: window.print();'>Print</a></td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<th>Equipment #</th>"
	Response.Write "<th>Description</th>"
	Response.Write "<th>Place checked</th>"
	Response.Write "<th>Date</th>"
	Response.Write "<th>dB level</th>"
	Response.Write "<th>In need of<br />Repair<br /><span style='font-size:6pt;font-weight:normal'>yes/no</span></th>"
	Response.Write "<th>Seal<br />condition<br /><span style='font-size:6pt;font-weight:normal'>good/leaking</span></th>"
	Response.Write "<th>Comments</th>"
	Response.Write "<th>Priority</th>"
	Response.Write "<th>Spare<br /><span style='font-size:6pt;font-weight:normal'>yes/no</span><th>"
	Response.Write "<th>&nbsp;</th>"
	Response.Write "</tr>"
	Response.Write "</thead>"

	Response.Write "<tfoot>"
	Response.Write "<tr><td colspan='11' style='width:100%;border:none'>&nbsp;"
	Response.Write "</td></tr>"
	Response.Write "</tfoot>"
	
	Response.Write "<tbody>"
	
	'Get the information for the equipment items and subitems.
	sqlString = "SELECT equipment_item_tag,equipment_item_description," & _
			"pm_priority,equipment_item_id " & _
			"FROM equipment_items " & _
			"WHERE equipment_type_id=6 " & _
			"ORDER BY equipment_item_id"
	Set rs = cn.Execute(sqlString)
	If Not rs.BOF Then
		rs.MoveFirst
		Do While Not rs.EOF
			'Fill in the variables.
			equipment_item_tag = rs("equipment_item_tag")
			equipment_item_description = rs("equipment_item_description")
			pm_priority = rs("pm_priority")
			itemID = rs("equipment_item_id")
			'Get the subitems for this item.
			sqlString = "SELECT equipment_items_subitem_description, " & _
			"test_front_bearing,test_rear_bearing,test_top " & _
			"FROM equipment_items_subitems " & _
			"WHERE equipment_item_id=" & itemID & _
			" ORDER BY equipment_items_subitem_id"
			rs2.Open sqlString,cn,3
			rs2.MoveLast
			rowCount = rs2.RecordCount
			If Not rs2.BOF Then
				rs2.MoveFirst
				count = 0
				Do While Not rs2.EOF
					count = count + 1
					equipment_items_subitem_description = rs2("equipment_items_subitem_description")
					test_front_bearing = CBool(rs2("test_front_bearing"))
					test_rear_bearing = CBool(rs2("test_rear_bearing"))
					test_top = CBool(rs2("test_top"))
					Response.Write "<tr>"
					If count = 1 Then
						Response.Write "<td rowspan='" & rowCount + 1 & "' style='text-align:center'>" & equipment_item_tag & "</td>"
					End If
					If count = 1 And test_front_bearing And test_rear_bearing Then
						Response.Write "<td rowspan='2'>" & equipment_items_subitem_description & "</td>"
					ElseIf count = 2 And test_top Then
						Response.Write "<td>" & equipment_items_subitem_description & "</td>"
					End If
					If count = 1 And test_front_bearing Then
						Response.Write "<td>front bearing</td>"
					ElseIf count = 2 And test_top Then
						Response.Write "<td>top</td>"
					End If
					If count = 1 Then
						Response.Write "<td rowspan='" & rowCount + 1 & "'>&nbsp;</td>"
					End If
					Response.Write "<td>&nbsp;</td>"
					If count = 1 Then
						Response.Write "<td rowspan='" & rowCount + 1 & "'>&nbsp;</td>"
						Response.Write "<td rowspan='" & rowCount + 1 & "'>&nbsp;</td>"
						Response.Write "<td rowspan='" & rowCount + 1 & "'>&nbsp;</td>"
						Response.Write "<td rowspan='" & rowCount + 1 & "' style='text-align:center'>" & pm_priority & "</td>"
					End If
					If count = 1 And test_front_bearing And test_rear_bearing Then
						Response.Write "<td rowspan='2'>&nbsp;</td>"
					ElseIf count = 2 And test_top Then
						Response.Write "<td>&nbsp;</td>"
					End If
					If count = 1 Then
						Response.Write "<td rowspan='" & rowCount + 1 & "' style='border:none'>" & equipment_item_description & "</td>"
					End If
					Response.Write "</tr>"
					If count = 1 And test_rear_bearing Then
						Response.Write "<tr>"
						Response.Write "<td>rear bearing</td>"
						Response.Write "<td>&nbsp;</td>"
						Response.Write "</tr>"
					End If
					
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			rs.MoveNext
		Loop
	End If
	rs.Close
	
	Set rs = Nothing
	Set rs2 = Nothing
	cn.Close
	Set cn = Nothing

	Response.Write "</tbody>"
	Response.Write "</table>"
	Response.Write "</form>"
	Response.Write "</body>"
	
Else
	Response.Write "<h1>You don't have permission to access this page.</h1>"
	Response.Write "<br />"
	Response.Write "<a href='" & request("http_referer") & "'>Return to previous page</a>"
End If
%>
</html>
