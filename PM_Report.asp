<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<meta http-equiv="X-UA-Compatible" content="IE=8" />
<html>
<head>
<script language="javascript">
</script>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>PM Report</title>
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
' Keith Brooks - Friday, January 13, 2012
'   Creation.
'*************

Dim sqlString
Dim cn
Dim rs
Dim rs2
Dim rs3
Dim rs4
Dim currentuser
Dim access
Dim rowCount
Dim count
Dim equipment_item_id
Dim equipment_item_tag
Dim equipment_item_description
Dim pm_priority
Dim equipment_items_subitem_description
Dim test_front_bearing
Dim test_rear_bearing
Dim test_top
Dim pm_date
Dim in_need_of_repair
Dim seal_condition
Dim comments
Dim front_bearing_db_level
Dim rear_bearing_db_level
Dim top_db_level
Dim spare1
Dim spare2
Dim startDate
Dim endDate
Dim id
Dim firstPass

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("inspections", "pm_report", currentuser)
If access <> "none" Then

	Response.Write "<body style='background-color:white'>"
		
	'Define the ado connection and recordset objects.
	set cn = CreateObject("adodb.connection")
	cn.CursorLocation = 3
	cn.Open = DBString
	set rs = CreateObject("adodb.recordset")
	Set rs2 = CreateObject("adodb.recordset")
	Set rs3 = CreateObject("adodb.recordset")
	Set rs4 = CreateObject("adodb.recordset")

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
	Response.Write "<td colspan='4' class='noprint' style='border:none;text-align:left;vertical-align:top;width:50%'><a href='pm_report_filter.asp'>Filter</a></td>"
	Response.Write "<td id='formtd' class='noprint' colspan='2' style='font-size:10pt;text-align:center'>"
	Response.Write "<a href='javascript: window.print();'>Print</a></td>"
	Response.Write "<td colspan='5' class='noprint' style='border:none;text-align:right;vertical-align:top;width:50%'><a href='' onclick='openhelp();return false;' title='Open the User Guide'>Help</a></td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td colspan='11' style='font-size:14pt;font-weight:bold;border:none;text-align:center'>Preventive Maintenance Report</td>"
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
	
	'Determine the start and end dates from the request objects.  If one or
	'more date was not specified, default to empty string.
	If IsDate(Request("start_date")) Then
		startDate = Request("start_date")
		Session("start_date") = startDate
	Else
		startDate = ""
	End If
	If IsDate(Request("end_date")) Then
		endDate = Request("end_date")
		Session("end_date") = endDate
	Else
		endDate = ""
	End If
	If IsNumeric(Request("equipment_item_id")) Then
		id = Request("equipment_item_id")
		Session("equipment_item_id") = id
	Else
		id = ""
	End If
	
	'Get the information for the equipment items.
	sqlString = "SELECT equipment_item_tag,equipment_item_description," & _
			"pm_priority,equipment_item_id " & _
			"FROM equipment_items " & _
			"WHERE equipment_type_id=6"
	If id <> "" Then
		sqlString = sqlString & " AND equipment_item_id=" & id
	End If
	sqlString = sqlString & " ORDER BY equipment_item_id"
	Set rs = cn.Execute(sqlString)
	If Not rs.BOF Then
		rs.MoveFirst
		Do While Not rs.EOF
			firstPass = True
			'Fill in the variables.
			equipment_item_tag = rs("equipment_item_tag")
			equipment_item_description = rs("equipment_item_description")
			pm_priority = rs("pm_priority")
			equipment_item_id = rs("equipment_item_id")
			
			'Get the pm data for this item for the specified dates.
			sqlString = "SELECT pm_date,in_need_of_repair,seal_condition," & _
				"comments,pm_data_id " & _
				"FROM pm_data " & _
				"WHERE equipment_item_id=" & equipment_item_id
			If IsDate(startDate) Then
				sqlString = sqlString & " AND pm_date>='" & FormatMySQLDate(startDate) & "'"
			End If
			If IsDate(endDate) Then
				sqlString = sqlString & " AND pm_date<='" & FormatMySQLDate(endDate) & "'"
			End If
			sqlString = sqlString & " ORDER BY pm_date"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				Do While Not rs2.EOF
				
					'Get the subitems for this item.
					sqlString = "SELECT equipment_items_subitem_description, " & _
						"test_front_bearing,test_rear_bearing,test_top, " & _
						"equipment_items_subitem_id " & _
						"FROM equipment_items_subitems " & _
						"WHERE equipment_item_id=" & equipment_item_id & _
						" ORDER BY equipment_items_subitem_id"
					rs3.Open sqlString,cn,3
					rs3.MoveLast
					rowCount = rs3.RecordCount
					If Not rs3.BOF Then
						rs3.MoveFirst
						count = 0
						Do While Not rs3.EOF
							count = count + 1
							equipment_items_subitem_description = rs3("equipment_items_subitem_description")
							test_front_bearing = CBool(rs3("test_front_bearing"))
							test_rear_bearing = CBool(rs3("test_rear_bearing"))
							test_top = CBool(rs3("test_top"))

							'Get the pm subitem data for the specified dates.  If a record exists,
							'fill the data into the variables; otherwise, initialize them.
							sqlString = "SELECT front_bearing_db_level,rear_bearing_db_level," & _
								"top_db_level,spare " & _
								"FROM pm_subitem_data " & _
								"WHERE pm_data_id=" & rs2("pm_data_id") & _
								" AND equipment_items_subitem_id=" & rs3("equipment_items_subitem_id")
							Set rs4 = cn.Execute(sqlString)
							If Not rs4.BOF Then
								rs4.MoveFirst
								
								'Loop through the data and display it.
								Do While Not rs4.EOF
									pm_date = rs2("pm_date")
									If IsNull(rs2("in_need_of_repair")) Then
										in_need_of_repair = 0
									Else
										in_need_of_repair = CBool(rs2("in_need_of_repair"))
									End If
									seal_condition = rs2("seal_condition")
									comments = rs2("comments")
									front_bearing_db_level = rs4("front_bearing_db_level")
									rear_bearing_db_level = rs4("rear_bearing_db_level")
									top_db_level = rs4("top_db_level")
									If count = 1 Then
										If IsNull(rs4("spare")) Then
											spare1 = 0
										Else
											spare1 = CBool(rs4("spare"))
										End If
									Else
										If IsNull(rs4("spare")) Then
											spare2 = 0
										Else
											spare2 = CBool(rs4("spare"))
										End If
									End If

									'Draw the row.
									Response.Write "<tr>"
									If count = 1 Then
										If firstPass Then
											Response.Write "<td rowspan='" & rowCount + 1 & "' style='text-align:center;border-bottom:0'>" & equipment_item_tag & "</td>"
										Else
											Response.Write "<td rowspan='" & rowCount + 1 & "' style='text-align:center;border-top:0'>&nbsp;</td>"
										End If
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
										Response.Write "<td rowspan='" & rowCount + 1 & "'  style='text-align:center'>" & pm_date & "</td>"
									End If
									If count = 1 And test_front_bearing Then
										Response.Write "<td>" & front_bearing_db_level & "</td>"
									ElseIf count = 2 And test_top Then
										Response.Write "<td>" & top_db_level & "</td>"
									End If
									If count = 1 Then
										Response.Write "<td rowspan='" & rowCount + 1 & "' style='text-align:center'>"
										If in_need_of_repair Then
											Response.Write "<input type='checkbox' id='in_need_of_repair' name='in_need_of_repair' value='1' checked disabled='disable' />"
										Else
											Response.Write "<input type='checkbox' id='in_need_of_repair' name='in_need_of_repair' value='1' disabled='disable' />"
										End If
										Response.Write "</td>"
										Response.Write "<td rowspan='" & rowCount + 1 & "'>" & seal_condition & "</td>"
										Response.Write "<td rowspan='" & rowCount + 1 & "'>" & comments & "</td>"
										Response.Write "<td rowspan='" & rowCount + 1 & "' style='text-align:center'>" & pm_priority & "</td>"
									End If
									If count = 1 And test_front_bearing And test_rear_bearing Then
										Response.Write "<td rowspan='2' style='text-align:center'>"
										If spare1 Then
											Response.Write "<input type='checkbox' id='spare1' name='spare1' value='1' checked disabled='disable' />"
										Else
											Response.Write "<input type='checkbox' id='spare1' name='spare1' value='1' disabled='disable' />"
										End If
										Response.Write "</td>"
									ElseIf count = 2 And test_top Then
										Response.Write "<td style='text-align:center'>"
										If spare2 Then
											Response.Write "<input type='checkbox' id='spare2' name='spare2' value='1' checked disabled='disable' />"
										Else
											Response.Write "<input type='checkbox' id='spare2' name='spare2' value='1' disabled='disable' />"
										End If
										Response.Write "</td>"
									End If
									If count = 1 Then
										If firstPass Then
											Response.Write "<td rowspan='" & rowCount & "' style='border:none'>" & equipment_item_description & "</td>"
										Else
											Response.Write "<td rowspan='" & rowCount & "' style='border:none'>&nbsp;</td>"
										End If
									Else
										Response.Write "<td style='text-align:left;border:none'>"
										Response.Write "&nbsp;"
										Response.Write "</td>"
									End If
									Response.Write "</tr>"
									If count = 1 And test_rear_bearing Then
										Response.Write "<tr>"
										Response.Write "<td>rear bearing</td>"
										Response.Write "<td>" & rear_bearing_db_level & "</td>"
										If rowCount = 1 Then
											Response.Write "<td style='text-align:left;border:none'>"
											Response.Write "&nbsp;"
											Response.Write "</td>"
										End If
										Response.Write "</tr>"
									End If
									rs4.MoveNext
									firstPass = False
								Loop
							End If
							rs4.Close
							rs3.MoveNext
						Loop
					End If
					rs3.Close
					rs2.MoveNext
				Loop
			Else
				Response.Write "<tr>"
				Response.Write "<td style='text-align:center'>" & equipment_item_tag & "</td>"
				Response.Write "<td colspan='9'>No PMs found</td>"
				Response.Write "<td style='border:none'>&nbsp;</td>"
				Response.Write "</tr>"
			End If
			rs2.Close
			rs.MoveNext
			If Not rs.EOF Then
				Response.Write "<tr>"
				Response.Write "<td colspan='10' style='border-left:none;border-right:none'>&nbsp;</td>"
				Response.Write "<td style='border:none'>&nbsp;</td>"
				Response.Write "</tr>"
			Else
				Response.Write "<tr>"
				Response.Write "<td colspan='10' style='border-left:none;border-right:none;border-bottom:none'>&nbsp;</td>"
				Response.Write "<td style='border:none'>&nbsp;</td>"
				Response.Write "</tr>"
			End If
		Loop
	End If
	rs.Close
	
	Set rs = Nothing
	Set rs2 = Nothing
	Set rs3 = Nothing
	Set rs4 = Nothing
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
