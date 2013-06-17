<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<html>
<head>
<script language="javascript">
</script>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>PSV Inspection</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="InspectionsFunctions.asp"-->
<link rel=STYLESHEET href='formstyle.css' type='text/css'>
</head>
<%
'*************
' Revision History
' 
' Keith Brooks - Monday, February 14, 2011
'   Creation.
'*************

Dim sqlString
Dim cn
Dim rs
Dim rs2
Dim currentuser
Dim access
Dim itemID
Dim equipment_item_tag
Dim equipment_item_name
Dim area
Dim assembly
Dim manufacturer
Dim model_number
Dim selected_orifice_size
Dim selected_orifice_size_units
Dim minimum_capacity
Dim minimum_capacity_units
Dim out_connect_size
Dim out_connect_size_units
Dim in_connect_size
Dim in_connect_size_units
Dim set_pressure
Dim set_pressure_units
Dim name_plate_stamp
Dim fluid_state
Dim fluid_service
Dim specification_number
Dim serial_number
Dim body_material
Dim trim_material
Dim noz_disk_material
Dim spring_material
Dim gasket_material
Dim bellows_material
Dim inspection_date
Dim valve_nameplate_matches
Dim valve_decontaminated
Dim inlet_piping_condition
Dim inlet_piping_required_cleaning
Dim outlet_piping_condition
Dim outlet_piping_required_cleaning
Dim field_initial_inspection_other
Dim condition_of_body
Dim inspection_company
Dim inspected_by
Dim inspected_date
Dim leaked_at_90pct_of_test_pr
Dim popped
Dim operated_properly
Dim returned_to_service
Dim test_conducted_by
Dim test_conducted_date
Dim test_only
Dim overhaul
Dim test_and_reset
Dim scrap_and_replace
Dim work_required_other
Dim disassembly_condition
Dim seats_lapped
Dim guide_replaced
Dim seats_replaced
Dim spring_replaced
Dim other_replaced
Dim other_repairs
Dim repair_company
Dim repaired_by
Dim repaired_date
Dim code_stamp
Dim authorization_number
Dim work_order_number
Dim next_inspection_due
Dim previous_inspection
Dim set_frequency
Dim set_frequency_units
Dim policy_insp_verify
Dim repair_required
Dim repair_type
Dim bubble_tight_at
Dim final_test_type
Dim final_test_set_press
Dim final_test_set_temp
Dim final_test_by
Dim final_test_date
Dim installed_by
Dim installed_date
Dim comment
Dim discrepency_comments
Dim discrepency_followup

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("inspections", "psv_manual_inspection", currentuser)
If access <> "none" Then

'	Response.Write "<body style='background-color:white' onload='window.print();window.close();'>"
	Response.Write "<body style='background-color:white'>"
		
	'Define the ado connection and recordset objects.
	set cn = CreateObject("adodb.connection")
	cn.Open = DBString
	set rs = CreateObject("adodb.recordset")

	Response.Write "<form id='form1' name='form1' action='inspectionaction.asp' method='post'>"
	
	If Request("itemID") <> "" Then
		'Get the static technical data for this item from the database.
		sqlString = "SELECT equipment_item_tag,equipment_item_name,assembly,area,b.* " & _
				"FROM equipment_items a INNER JOIN psv_technical_data b " & _
				"ON a.equipment_item_id=b.equipment_item_id " & _
				"WHERE a.equipment_item_id=" & Request("itemID")
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			'Fill in the static technical data variables.
			rs.MoveFirst
			equipment_item_tag = rs("equipment_item_tag")
			equipment_item_name = rs("equipment_item_name")
			area = rs("area")
			assembly = rs("assembly")
			manufacturer = rs("manufacturer")
			model_number = rs("model_number")
			selected_orifice_size = rs("selected_orifice_size")
			selected_orifice_size_units = rs("selected_orifice_size_units")
			minimum_capacity = rs("minimum_capacity")
			minimum_capacity_units = rs("minimum_capacity_units")
			out_connect_size = rs("out_connect_size")
			out_connect_size_units = rs("out_connect_size_units")
			in_connect_size = rs("in_connect_size")
			in_connect_size_units = rs("in_connect_size_units")
			set_pressure = rs("set_pressure")
			set_pressure_units = rs("set_pressure_units")
			name_plate_stamp = rs("name_plate_stamp")
			fluid_state = rs("fluid_state")
			fluid_service = rs("fluid_service")
			specification_number = rs("specification_number")
			serial_number = rs("serial_number")
			body_material = rs("body_material")
			trim_material = rs("trim_material")
			noz_disk_material = rs("noz_disk_material")
			spring_material = rs("spring_material")
			gasket_material = rs("gasket_material")
			bellows_material = rs("bellows_material")
		End If
		rs.Close
			
		'Determine the previous and next inspection dates.
		sqlString = "SELECT inspection_date,set_frequency,set_frequency_units " & _
				"FROM psv_inspection_data " & _
				"WHERE inspection_date=" & _
				"(SELECT MAX(inspection_date) FROM psv_inspection_data " & _
				"WHERE equipment_item_id=" & Request("itemID") & ")"
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			If Not IsNull(rs("inspection_date")) Then
				previous_inspection = FormatDateTime(rs("inspection_date"),2)
			Else
				previous_inspection = ""
			End If
			If Not IsNull(rs("set_frequency")) Then
				set_frequency = rs("set_frequency")
			Else
				set_frequency = 0
			End If
			If Not IsNull(rs("set_frequency_units")) Then
				set_frequency_units = rs("set_frequency_units")
			Else
				set_frequency_units = ""
			End If
		Else
			previous_inspection = 0
			set_frequency = 0
			set_frequency_units = ""
		End If
		rs.Close
		'If set frequency or its units are not specified, look the default values
		'up in the equipment types table.
		If set_frequency = 0 Or set_frequency_units = "" Then
			sqlString = "SELECT inspection_interval,inspection_interval_units " & _
					"FROM equipment_types WHERE equipment_type_name='PSV'"
			Set rs = cn.Execute(sqlString)
			If Not rs.BOF Then
				rs.MoveFirst
				If Not IsNull(rs("inspection_interval")) And set_frequency = 0 Then
					set_frequency = rs("inspection_interval")
				End If
				If Not IsNull(rs("inspection_interval_units")) And set_frequency_units = "" Then
					set_frequency_units = rs("inspection_interval_units")
				End If
			End If
			rs.Close
		End If
		If CInt(set_frequency) > 0 And set_frequency_units <> "" Then
			Dim units
			Select Case LCase(set_frequency_units)
				Case "years","year","yyyy","yy","y"
					units = "yyyy"
				Case "months","month","mon","m"
					units = "m"
				Case "days","day","dd","d"
					units = "d"
				Case Else
					units = ""
			End Select
			next_inspection_due = DateAdd(units,set_frequency,Date)
		Else
			next_inspection_due = ""
		End If
	End If
	
	Response.Write "<table border='0' align='center' width='100%'>"
	Response.Write "<thead>"
	Response.Write "<tr><td>"
	'Draw the header.
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' class='noprint' colspan='2' style='font-size:10pt;text-align:center'>"
	Response.Write "<a href='javascript: window.print();'>Print</a></td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td style='text-align:left;vertical-align:top;font-size:12pt;font-weight:bold'>Relief Valve - Inspection Report</td>"
	Response.Write "<td style='text-align:right;font-size:14pt;font-weight:bold'>AKSA</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td colspan='2' style='border-top:1px solid black;border-bottom:1px solid black;height:2px;padding:0px;line-height:2px'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	'Draw the form.
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:15%'>Inspection Date: </td>"
	Response.Write "<td id='blanktd' style='width:25%'>&nbsp;</td>"
	Response.Write "<td id='formtd' style='width:10%'>Area: </td>"
	Response.Write "<td id='formtd' style='width:10%'>" & area & "</td>"
	Response.Write "<td id='formtd' style='width:10%'>Equip No: </td>"
	Response.Write "<td id='formtd' style='width:30%'>" & equipment_item_tag & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Equip Name: </td>"
	Response.Write "<td id='formtd' colspan='3'>" & equipment_item_name & "</td>"
	Response.Write "<td id='formtd'>Assembly: </td>"
	Response.Write "<td id='formtd'>" & assembly & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='6' style='border-bottom:1px solid black;height:5px;line-height:5px;padding:0px'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "</td></tr>"
	Response.Write "</thead>"

	Response.Write "<tfoot>"
	Response.Write "<tr><td style='width:100%'>&nbsp;"
	Response.Write "</td></tr>"
	Response.Write "</tfoot>"
	
	Response.Write "<tbody><tr><td>"
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='5'>RELIEF VALVE DESIGN:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='techtd' style='width:20%'>Manufacturer: </td>"
	Response.Write "<td id='techtd' colspan='2'>" & manufacturer & "</td>"
	Response.Write "<td id='techtd' style='width:17%'>Spec Number: </td>"
	Response.Write "<td id='techtd' style='width:33%'>" & specification_number & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='techtd'>Model Number: </td>"
	Response.Write "<td id='techtd' colspan='2'>" & model_number & "</td>"
	Response.Write "<td id='techtd'>Serial Number: </td>"
	Response.Write "<td id='techtd'>" & serial_number & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='techtd'>Selected Orifice Size: </td>"
	Response.Write "<td id='techtd' style='width:10%'>" & selected_orifice_size & "</td>"
	Response.Write "<td id='techtd' style='width:20%'>" & selected_orifice_size_units & "</td>"
	Response.Write "<td id='techtd'>Body Material: </td>"
	Response.Write "<td id='techtd'>" & body_material & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='techtd'>Capacity: </td>"
	Response.Write "<td id='techtd'>" & minimum_capacity & "</td>"
	Response.Write "<td id='techtd'>" & minimum_capacity_units & "</td>"
	Response.Write "<td id='techtd'>Trim Material: </td>"
	Response.Write "<td id='techtd'>" & trim_material & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='techtd'>Out Connect: </td>"
	Response.Write "<td id='techtd'>" & out_connect_size & "</td>"
	Response.Write "<td id='techtd'>" & out_connect_size_units & "</td>"
	Response.Write "<td id='techtd'>Noz/Disk Material: </td>"
	Response.Write "<td id='techtd'>" & noz_disk_material & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='techtd'>In Connect: </td>"
	Response.Write "<td id='techtd'>" & in_connect_size & "</td>"
	Response.Write "<td id='techtd'>" & in_connect_size_units & "</td>"
	Response.Write "<td id='techtd'>Spring Material: </td>"
	Response.Write "<td id='techtd'>" & spring_material & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='techtd'>Set Pressure: </td>"
	Response.Write "<td id='techtd'>" & set_pressure & "</td>"
	Response.Write "<td id='techtd'>" & set_pressure_units & "</td>"
	Response.Write "<td id='techtd'>Gasket Material: </td>"
	Response.Write "<td id='techtd'>" & gasket_material & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='techtd'>Name Plate: </td>"
	Response.Write "<td id='techtd' colspan='2'>" & name_plate_stamp & "</td>"
	Response.Write "<td id='techtd'>Bellows Material: </td>"
	Response.Write "<td id='techtd'>" & bellows_material & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='techtd'>Fluid State: </td>"
	Response.Write "<td id='techtd' colspan='2'>" & fluid_state & "</td>"
	Response.Write "<td id='techtd'>&nbsp;</td>"
	Response.Write "<td id='techtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='techtd'>Fluid Service: </td>"
	Response.Write "<td id='techtd' colspan='2'>" & fluid_service & "</td>"
	Response.Write "<td id='techtd'>&nbsp;</td>"
	Response.Write "<td id='techtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "</table>"

	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='5'>FIELD/INITIAL INSPECTION:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:40%'>Valve Nameplate Data Matches Above Data:</td>"
	Response.Write "<td id='formtd' style='width:10%'>"
	Response.Write "<input type='checkbox' class='checkbox' id='valve_nameplate_matches' name='valve_nameplate_matches' value='1' /></td>"
	Response.Write "<td id='formtd' style='width:20%'>Valve Decontaminated:</td>"
	Response.Write "<td id='formtd' style='width:30%'>"
	Response.Write "<input type='checkbox' class='checkbox' id='valve_decontaminated' name='valve_decontaminated' value='1' /></td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='3' style='text-align:center'>INLET PIPING</td>"
	Response.Write "<td id='grouptd' colspan='3' style='text-align:center'>OUTLET PIPING</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:26%'>Condition:</td>"
	Response.Write "<td id='blanktd' style='width:23%'>&nbsp;</td>"
	Response.Write "<td id='formtd' style='width:1%'>&nbsp;</td>"
	Response.Write "<td id='formtd' style='width:20%'>Condition:</td>"
	Response.Write "<td id='blanktd' style='width:25%'>&nbsp;</td>"
	Response.Write "<td id='formtd' style='width:5%'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Required Cleaning:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='inlet_piping_required_cleaning' name='inlet_piping_required_cleaning' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Required Cleaning:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='outlet_piping_required_cleaning' name='outlet_piping_required_cleaning' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Other (Specify):</td>"
	Response.Write "<td id='blanktd' colspan='4'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='blanktd' colspan='4'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Condition of Body (Specify):</td>"
	Response.Write "<td id='blanktd' colspan='4'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr><td id='formtd' colspan='5'>&nbsp;</td></tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='font-weight:bold'>Inspection Company:</td>"
	Response.Write "<td id='blanktd' colspan='4'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='font-weight:bold'>Inspected By:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd' style='font-weight:bold'>Date: </td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='6'>&nbsp;</td>"
	Response.Write "</tr>"
	
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='6'>INITIAL TEST:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Leaked at 90% of Test Pr.:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Popped: </td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Operated Properly:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='operated_properly' name='operated_properly' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Returned To Service:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='returned_to_service' name='returned_to_service' value='1' /></td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Test Conducted By:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Date:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"

	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='6'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='6'>WORK REQUIRED:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Test Only: </td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='test_only' name='test_only' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Overhaul: </td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='overhaul' name='overhaul' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Test and Reset: </td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='test_and_reset' name='test_and_reset' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Scrap and Replace: </td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='scrap_and_replace' name='scrap_and_replace' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Other (Specify): </td>"
	Response.Write "<td id='blanktd' colspan='4'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"

	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='6'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='6'>REPAIRS MADE:</td>"
	Response.Write "</tr>"
	Response.Write "<td id='formtd'>Disassembly Condition: </td>"
	Response.Write "<td id='blanktd' colspan='4'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Seats Lapped: </td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='seats_lapped' name='seats_lapped' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Guide Replaced: </td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='guide_replaced' name='guide_replaced' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Seats Replaced: </td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='seats_replaced' name='seats_replaced' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Spring Replaced: </td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='spring_replaced' name='spring_replaced' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Other Replaced (Specify): </td>"
	Response.Write "<td id='blanktd' colspan='4'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Other Repairs: </td>"
	Response.Write "<td id='blanktd' colspan='4'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Repair Company: </td>"
	Response.Write "<td id='blanktd' colspan='4'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Repaired By: </td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Date: </td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Code Stamp: </td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Authorization No: </td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='6'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='3' style='padding:0px'>"
	Response.Write "<table style='width:100%;border:1px solid black'>"
	Response.Write "<tr>"
	Response.Write "<td id='smalltd' style='width:40%'>Policy Insp. Verify:</td>"
	Response.Write "<td id='smalltd' style='width:20%'>"
	Response.Write "<input type='radio' class='radio' id='policy_insp_verify' name='policy_insp_verify' value='yes' />yes</td>"
	Response.Write "<td id='smalltd' style='width:20%'>"
	Response.Write "<input type='radio' class='radio' id='policy_insp_verify' name='policy_insp_verify' value='no' />no</td>"
	Response.Write "<td id='smalltd' style='width:20%'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='smalltd'>Repair Required:</td>"
	Response.Write "<td id='smalltd'>"
	Response.Write "<input type='radio' class='radio' id='repair_required' name='repair_required' value='yes' />yes</td>"
	Response.Write "<td id='smalltd'>"
	Response.Write "<input type='radio' class='radio' id='repair_required' name='repair_required' value='no' />no</td>"
	Response.Write "<td id='smalltd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='smalltd'>Repair Type:</td>"
	Response.Write "<td id='smalltd'>"
	Response.Write "<input type='radio' class='radio' id='repair_type' name='repair_type' value='none' />none</td>"
	Response.Write "<td id='smalltd'>"
	Response.Write "<input type='radio' class='radio' id='repair_type' name='repair_type' value='major' />major</td>"
	Response.Write "<td id='smalltd'>"
	Response.Write "<input type='radio' class='radio' id='repair_type' name='repair_type' value='minor' />minor</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "</td>"
	Response.Write "<td id='formtd' colspan='3' style='padding:0px 0px 0px 5px'>"
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:39%'>Work Order Number:</td>"
	Response.Write "<td id='blanktd' colspan='3'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Next Inspection Due:</td>"
	Response.Write "<td id='formtd' colspan='4'>" & next_inspection_due & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Previous Inspection:</td>"
	Response.Write "<td id='formtd' style='width:21%'>" & previous_inspection & "</td>"
	Response.Write "<td id='formtd' style='width:20%'>Set Freq:</td>"
	Response.Write "<td id='formtd' style='width:9%'>" & set_frequency & "</td>"
	Response.Write "<td id='formtd' style='width:10%'>" & set_frequency_units & "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	
	Response.Write "<div style='page-break-before:always;font-size:1;margin:0;border:0'><span style='visibility:hidden'>-</span></div>"
	
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='8'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='8'>FINAL TEST:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Bubble Tight At:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd' colspan='3'>PSIG</td>"
	Response.Write "<td id='formtd' style='width:17%'>Type:</td>"
	Response.Write "<td id='blanktd' style='width:27%'>&nbsp;</td>"
	Response.Write "<td id='formtd' style='width:5%'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:20%'>Set: </td>"
	Response.Write "<td id='blanktd' style='width:9%'>&nbsp;</td>"
	Response.Write "<td id='formtd' style='width:8%'>PSIG at</td>"
	Response.Write "<td id='blanktd' style='width:9%'>&nbsp;</td>"
	Response.Write "<td id='formtd' style='width:4%'>F</td>"
	Response.Write "<td id='formtd' colspan='3'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Final Test By: </td>"
	Response.Write "<td id='blanktd' colspan='3'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Date: </td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"

	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='8'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='8'>INSTALLATION:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Installed By: </td>"
	Response.Write "<td id='blanktd' colspan='3'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Date: </td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='8'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='8'>COMMENT:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd' colspan='7'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd' colspan='7'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"

	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='8'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='8'>DISCREPENCY:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='8' style='font-weight:bold'>&nbsp;&nbsp;COMMENTS:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd' colspan='7'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd' colspan='7'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='8' style='font-weight:bold'>&nbsp;&nbsp;FOLLOW-UP:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd' colspan='7'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd' colspan='7'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	
	Set rs = Nothing
	cn.Close
	Set cn = Nothing

	Response.Write "</td></tr></tbody>"
	
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
