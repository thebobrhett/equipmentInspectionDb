<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<html>
<head>
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
' Keith Brooks - Wednesday, March 16, 2011
'   Creation.
'*************

Dim sqlString
Dim cn
Dim rs
Dim rs2
Dim criteria
Dim currentuser
Dim access
Dim editMode
Dim itemID
'Read-only items
Dim equipment_item_tag
Dim equipment_item_name
Dim area
Dim assembly
Dim manufacturer
Dim model_number
Dim vacuum_set_point
Dim vacuum_set_point_units
Dim pressure_set_point
Dim pressure_set_point_units
Dim pad_set_point
Dim pad_set_point_units
Dim reg_gauge_range_from
Dim reg_gauge_range_to
Dim reg_gauge_range_units
Dim fluid_service
Dim specification_number
Dim serial_number
Dim arrester_manufacturer
Dim arrester_model_number
Dim arrester_serial_number
Dim arrester_spec_number
Dim fluid_state
Dim previous_inspection
Dim next_inspection_due
Dim set_frequency
Dim set_frequency_units

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
	
	'Save the equipment type for use by the action page.
	Response.Write "<input type='hidden' id='equipType' name='equipType' value='psv' />"
	
	If Request("itemID") <> "" Then
		'Save the value of the item id.
		Response.Write "<input type='hidden' id='itemID' name='itemID' value='" & Request("itemID") & "' />"

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
			vacuum_set_point = rs("vacuum_set_point")
			vacuum_set_point_units = rs("vacuum_set_point_units")
			pressure_set_point = rs("pressure_set_point")
			pressure_set_point_units = rs("pressure_set_point_units")
			pad_set_point = rs("pad_set_point")
			pad_set_point_units = rs("pad_set_point_units")
			reg_gauge_range_from = rs("reg_gauge_range_from")
			reg_gauge_range_to = rs("reg_gauge_range_to")
			reg_gauge_range_units = rs("reg_gauge_range_units")
			fluid_service = rs("fluid_service")
			specification_number = rs("specification_number")
			serial_number = rs("serial_number")
			arrester_manufacturer = rs("arrester_manufacturer")
			arrester_model_number = rs("arrester_model_number")
			arrester_serial_number = rs("arrester_serial_number")
			arrester_spec_number = rs("arrester_spec_number")
			fluid_state = rs("fluid_state")
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
				previous_inspection = "NONE"
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
			previous_inspection = "NONE"
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
	Response.Write "<td style='text-align:left;vertical-align:top;font-size:12pt;font-weight:bold'>Conservation Vent - Inspection Report</td>"
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
	Response.Write "<td id='grouptd' colspan='5'>CONSERVATION VENT DESIGN:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='techtd' style='width:24%'>Manufacturer: </td>"
	Response.Write "<td id='techtd' colspan='2'>" & manufacturer & "</td>"
	Response.Write "<td id='techtd' style='width:20%'>Specification Number: </td>"
	Response.Write "<td id='techtd' style='width:30%'>" & specification_number & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='techtd'>Model Number: </td>"
	Response.Write "<td id='techtd' colspan='2'>" & model_number & "</td>"
	Response.Write "<td id='techtd'>Serial Number: </td>"
	Response.Write "<td id='techtd'>" & serial_number & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='techtd'>Vacuum Set Point: </td>"
	Response.Write "<td id='techtd' style='width:10%'>" & vacuum_set_point & "</td>"
	Response.Write "<td id='techtd' style='width:16%'>" & vacuum_set_point_units & "</td>"
	Response.Write "<td id='techtd'>Fl. Arrester Mfgr: </td>"
	Response.Write "<td id='techtd'>" & arrester_manufacturer & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='techtd'>Pressure Set Point: </td>"
	Response.Write "<td id='techtd'>" & pressure_set_point & "</td>"
	Response.Write "<td id='techtd'>" & pressure_set_point_units & "</td>"
	Response.Write "<td id='techtd'>Fl. Arrester Model: </td>"
	Response.Write "<td id='techtd'>" & arrester_model_number & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='techtd'>N2 Pad Reg Set Point: </td>"
	Response.Write "<td id='techtd'>" & pad_set_point & "</td>"
	Response.Write "<td id='techtd'>" & pad_set_point_units & "</td>"
	Response.Write "<td id='techtd'>Fl. Arrester Serial: </td>"
	Response.Write "<td id='techtd'>" & arrester_serial_number & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='techtd'>N2 Pad Reg Gauge Range: </td>"
	Response.Write "<td id='techtd'>" & reg_gauge_range_from & "</td>"
	Response.Write "<td id='techtd'>to&nbsp;&nbsp;" & reg_gauge_range_to & "</td>"
	Response.Write "<td id='techtd'>Fl. Arrester Spec: </td>"
	Response.Write "<td id='techtd'>" & arrester_spec_number & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='techtd'>Fluid Service: </td>"
	Response.Write "<td id='techtd' colspan='2'>" & fluid_service & "</td>"
	Response.Write "<td id='techtd'>Fluid State: </td>"
	Response.Write "<td id='techtd'>" & fluid_state & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr><td id='techtd' colspan='5'>&nbsp;</td></tr>"
	Response.Write "</table>"

	'Draw the data entry section.
	Response.Write "<table style='width:100%;border:none'>"	
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='4'>INSPECTION:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:33%'>Vent Nameplate Data Match Above Data: </td>"
	Response.Write "<td id='formtd' style='width:17%'>"
	Response.Write "<input type='checkbox' class='checkbox' id='vent_nameplate_matches' name='vent_nameplate_matches' value='1' /></td>"
	Response.Write "<td id='formtd' style='width:23%'>Vent Decontaminated: </td>"
	Response.Write "<td id='formtd' style='width:27%'>"
	Response.Write "<input type='checkbox' class='checkbox' id='vent_decontaminated' name='vent_decontaminated' value='1' /></td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Flame Arr Nameplate Match Above Data: </td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='flame_arr_nameplate_matches' name='flame_arr_nameplate_matches' value='1' /></td>"
	Response.Write "<td id='formtd'>Flame Arr Decontaminated: </td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='flame_arr_decontaminated' name='flame_arr_decontaminated' value='1' /></td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='3' style='text-align:center'>VENT INLET</td>"
	Response.Write "<td id='grouptd' colspan='5' style='text-align:center'>VENT PIPING</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:24%'>Condition: </td>"
	Response.Write "<td id='blanktd' style='width:25%'>&nbsp;</td>"
	Response.Write "<td id='formtd' style='width:1%'>&nbsp;</td>"
	Response.Write "<td id='formtd' style='width:20%'>Condition: </td>"
	Response.Write "<td id='blanktd' colspan='3'>&nbsp;</td>"
	Response.Write "<td id='formtd' style='width:5%'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Requires Cleaning: </td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='vent_inlet_requires_cleaning' name='vent_inlet_requires_cleaning' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Requires Cleaning: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	Response.Write "<input type='checkbox' class='checkbox' id='vent_piping_requires_cleaning' name='vent_piping_requires_cleaning' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='vertical-align:top'>Other: </td>"
	Response.Write "<td id='blanktd' colspan='6'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='blanktd' colspan='6'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Body: </td>"
	Response.Write "<td id='blanktd' colspan='6'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='3' style='text-align:center'>FLAME ARRESTER/SCREEN</td>"
	Response.Write "<td id='grouptd' colspan='5' style='text-align:center'>PADDING REGULATOR</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Condition: </td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Condition: </td>"
	Response.Write "<td id='blanktd' colspan='3'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Requires Cleaning:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='flame_arrester_requires_cleaning' name='flame_arrester_requires_cleaning' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Gauge Condition:</td>"
	Response.Write "<td id='blanktd' colspan='3'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Requires Repair:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='flame_arrester_requires_repair' name='flame_arrester_requires_repair' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Replace Regulator:</td>"
	Response.Write "<td id='formtd' style='width:5%'>"
	Response.Write "<input type='checkbox' class='checkbox' id='replace_regulator' name='replace_regulator' value='1' /></td>"
	Response.Write "<td id='formtd' style='width:15%'>Replace Gauge:</td>"
	Response.Write "<td id='formtd' style='width:5%'>"
	Response.Write "<input type='checkbox' class='checkbox' id='replace_gauge' name='replace_gauge' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='3' style='text-align:center'>PRESSURE PALLET</td>"
	Response.Write "<td id='grouptd' colspan='5' style='text-align:center'>VACUUM PALLET</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Condition: </td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Condition: </td>"
	Response.Write "<td id='blanktd' colspan='3'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Requires Cleaning: </td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='pressure_pallet_requires_cleaning' name='pressure_pallet_requires_cleaning' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Requires Cleaning: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	Response.Write "<input type='checkbox' class='checkbox' id='vacuum_pallet_requires_cleaning' name='vacuum_pallet_requires_cleaning' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Requires Repair: </td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='pressure_pallet_requires_repair' name='pressure_pallet_requires_repair' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Requires Repair: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	Response.Write "<input type='checkbox' class='checkbox' id='vacuum_pallet_requires_repair' name='vacuum_pallet_requires_repair' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Operated Manually: </td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='pressure_pallet_operated_manually' name='pressure_pallet_operated_manually' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Operated Manually: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	Response.Write "<input type='checkbox' class='checkbox' id='vacuum_pallet_operated_manually' name='vacuum_pallet_operated_manually' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr><td id='formtd' colspan='8'>&nbsp;</td></tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='font-weight:bold'>Inspection Company: </td>"
	Response.Write "<td id='blanktd' colspan='6'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='font-weight:bold'>Field Inspection By: </td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd' style='font-weight:bold'>Date: </td>"
	Response.Write "<td id='blanktd' colspan='3'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr><td id='formtd' colspan='8'>&nbsp;</td></tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='8'>REPAIRS MADE:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='3' style='text-align:center'>PRESSURE PALLET</td>"
	Response.Write "<td id='grouptd' colspan='5' style='text-align:center'>VACUUM PALLET</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Cleaned: </td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='pressure_pallet_cleaned' name='pressure_pallet_cleaned' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Cleaned: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	Response.Write "<input type='checkbox' class='checkbox' id='vacuum_pallet_cleaned' name='vacuum_pallet_cleaned' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Seats Replaced: </td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='pressure_pallet_seats_replaced' name='pressure_pallet_seats_replaced' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Seats Replaced: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	Response.Write "<input type='checkbox' class='checkbox' id='vacuum_pallet_seats_replaced' name='vacuum_pallet_seats_replaced' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Guides Replaced: </td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='pressure_pallet_guides_replaced' name='pressure_pallet_guides_replaced' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Guides Replaced: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	Response.Write "<input type='checkbox' class='checkbox' id='vacuum_pallet_guides_replaced' name='vacuum_pallet_guides_replaced' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Other: </td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Other: </td>"
	Response.Write "<td id='blanktd' colspan='3'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='blanktd' colspan='3'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:30%'>Flame Arrester Screen Cleaned: </td>"
	Response.Write "<td id='formtd' style='width:20%'>"
	Response.Write "<input type='checkbox' class='checkbox' id='flame_arrester_screen_cleaned' name='flame_arrester_screen_cleaned' value='1' /></td>"
	Response.Write "<td id='formtd' colspan='3'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Conservation Vent Replaced:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='conservation_vent_replaced' name='conservation_vent_replaced' value='1' /></td>"
	Response.Write "<td id='formtd' style='width:20%'>Serial No. New Vent:</td>"
	Response.Write "<td id='blanktd' style='width:25%'>&nbsp;</td>"
	Response.Write "<td id='formtd' style='width:5%'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Flame Arrester Replaced:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='flame_arrester_replaced' name='flame_arrester_replaced' value='1' /></td>"
	Response.Write "<td id='formtd'>Serial No. FA:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='5'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2' style='padding:0px'>"
	Response.Write "<table style='width:100%;border:1px solid black'>"
	Response.Write "<tr>"
	Response.Write "<td id='smalltd' style='width:40%'>Policy Insp. Verify:</td>"
	Response.Write "<td id='smalltd' style='width:5%'>"
	Response.Write "<input type='radio' class='radio' id='policy_insp_verify' name='policy_insp_verify' value='yes' />yes</td>"
	Response.Write "<td id='smalltd' style='width:5%'>"
	Response.Write "<input type='radio' class='radio' id='policy_insp_verify' name='policy_insp_verify' value='no' />no</td>"
	Response.Write "<td id='smalltd' style='width:40%'>Repair Performed:</td>"
	Response.Write "<td id='smalltd' style='width:5%'>"
	Response.Write "<input type='radio' class='radio' id='repair_performed' name='repair_performed' value='yes' />yes</td>"
	Response.Write "<td id='smalltd' style='width:5%'>"
	Response.Write "<input type='radio' class='radio' id='repair_performed' name='repair_performed' value='no' />no</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='smalltd'>Repair Required:</td>"
	Response.Write "<td id='smalltd'>"
	Response.Write "<input type='radio' class='radio' id='repair_required' name='repair_required' value='yes' />yes</td>"
	Response.Write "<td id='smalltd'>"
	Response.Write "<input type='radio' class='radio' id='repair_required' name='repair_required' value='no' />no</td>"
	Response.Write "<td id='smalltd' colspan='3'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='smalltd'>Repair Type:</td>"
	Response.Write "<td id='smalltd'>"
	Response.Write "<input type='radio' class='radio' id='repair_type' name='repair_type' value='none' />none</td>"
	Response.Write "<td id='smalltd'>"
	Response.Write "<input type='radio' class='radio' id='repair_type' name='repair_type' value='major' />major</td>"
	Response.Write "<td id='smalltd'>"
	Response.Write "<input type='radio' class='radio' id='repair_type' name='repair_type' value='minor' />minor</td>"
	Response.Write "<td id='smalltd' colspan='2'>&nbsp;</td>"
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
	
	Response.Write "<br />"
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='10'>REPAIRS MADE (cont):</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:24%'>Regulator Repaired:</td>"
	Response.Write "<td id='formtd' style='width:25%'>"
	Response.Write "<input type='checkbox' class='checkbox' id='regulator_repaired' name='regulator_repaired' value='1' /></td>"
	Response.Write "<td id='formtd' style='width:1%'>&nbsp;</td>"
	Response.Write "<td id='formtd' style='width:20%'>Set Point:</td>"
	Response.Write "<td id='blanktd' colspan='3'>&nbsp;</td>"
	Response.Write "<td id='formtd' style='width:3%'>(units)</td>"
	Response.Write "<td id='blanktd' style='width:7%'>&nbsp</td>"
	Response.Write "<td id='formtd' style='width:5%'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Regulator Gauge Repaired:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='regulator_gauge_repaired' name='regulator_gauge_repaired' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Range:</td>"
	Response.Write "<td id='blanktd' style='width:7%'>&nbsp;</td>"
	Response.Write "<td id='formtd' style='width:1%;text-align:center'>to</td>"
	Response.Write "<td id='blanktd' style='width:7%'>&nbsp;</td>"
	Response.Write "<td id='formtd'>(units)</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Repair Company: </td>"
	Response.Write "<td id='blanktd' colspan='8'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Repaired By:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Date:</td>"
	Response.Write "<td id='blanktd' colspan='5'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Cleaned By:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Date:</td>"
	Response.Write "<td id='blanktd' colspan='5'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr><td id='formtd' colspan='10'>&nbsp;</td></tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='10'>INSTALLATION:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Flange Bolts Replaced:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='flange_bolts_replaced' name='flange_bolts_replaced' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Type:</td>"
	Response.Write "<td id='blanktd' colspan='5'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Flange Bolts Torqued:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='flange_bolts_torqued' name='flange_bolts_torqued' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Type:</td>"
	Response.Write "<td id='blanktd' colspan='5'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Regulator Replaced:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='regulator_replaced' name='regulator_replaced' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Set Point:</td>"
	Response.Write "<td id='blanktd' colspan='3'>&nbsp;</td>"
	Response.Write "<td id='formtd'>(units)</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Regulator Gauge Replaced:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='regulator_gauge_replaced' name='regulator_gauge_replaced' value='1' /></td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Range:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd' style='text-align:center'>to</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>(units)</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Installed By:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>Date:</td>"
	Response.Write "<td id='blanktd' colspan='5'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='10'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='10'>COMMENT:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd' colspan='9'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd' colspan='9'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"

	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='10'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='10'>DISCREPANCY:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='10' style='font-weight:bold'>&nbsp;&nbsp;COMMENTS:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd' colspan='9'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd' colspan='9'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='10' style='font-weight:bold'>&nbsp;&nbsp;FOLLOW-UP:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd' colspan='9'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd' colspan='9'>&nbsp;</td>"
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
