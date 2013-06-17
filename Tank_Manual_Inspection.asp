<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<html>
<head>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Vessel Inspection</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="InspectionsFunctions.asp"-->
<link rel=STYLESHEET href='formstyle.css' type='text/css'>
</head>
<%
'*************
' Revision History
' 
' Keith Brooks - Thursday, March 4, 2011
'   Creation.
'*************

Dim sqlString
Dim cn
Dim rs
Dim rs2
Dim criteria
Dim currentuser
Dim access
Dim itemID
Dim equipment_item_tag
Dim equipment_item_name
Dim area
Dim assembly
Dim state_number
Dim relief_device
Dim relief_device_pressure
Dim relief_device_pressure_units
Dim lethal_service
Dim capacity
Dim capacity_units
Dim weight_empty
Dim weight_empty_units
Dim height_length
Dim height_length_units
Dim inside_diameter
Dim inside_diameter_units
Dim shell_material
Dim shell_thickness
Dim shell_thickness_units
Dim shell_min_thickness
Dim shell_min_thickness_units
Dim head_material
Dim head_thickness
Dim head_thickness_units
Dim head_min_thickness
Dim head_min_thickness_units
Dim lining_material
Dim lining_thickness
Dim lining_thickness_units
Dim jacket_material
Dim jacket_thickness
Dim jacket_thickness_units
Dim mawp
Dim mawp_units
Dim shell_test_press
Dim shell_test_press_units
Dim jacket_test_press
Dim jacket_test_press_units
Dim date_built
Dim national_board_number
Dim lining_mfgr
Dim manufacturer
Dim mfgr_serial_number
Dim drawing_number
Dim jacket_type_description
Dim inspection_date
Dim next_inspection_due
Dim previous_inspection_date
Dim set_frequency
Dim set_frequency_units

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("inspections", "tank_manual_inspection", currentuser)
If access <> "none" Then

'	Response.Write "<body style='background-color:white' onload='window.print();window.close();'>"
	Response.Write "<body style='background-color:white'>"
		
	'Define the ado connection and recordset objects.
	set cn = CreateObject("adodb.connection")
	cn.Open = DBString
	set rs = CreateObject("adodb.recordset")

	Response.Write "<form id='form1' name='form1' action='inspectionaction.asp' method='post'>"
	
	If Request("itemID") <> "" Then
		'Save the value of the item id.
		Response.Write "<input type='hidden' id='itemID' name='itemID' value='" & Request("itemID") & "' />"

		'Get the static technical data for this item from the database.
		sqlString = "SELECT equipment_item_tag,equipment_item_name,assembly,area,b.* " & _
				"FROM equipment_items a INNER JOIN tank_technical_data b " & _
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
			state_number = rs("state_number")
			relief_device = rs("relief_device")
			relief_device_pressure = rs("relief_device_pressure")
			relief_device_pressure_units = rs("relief_device_pressure_units")
			lethal_service = rs("lethal_service")
			capacity = rs("capacity")
			capacity_units = rs("capacity_units")
			weight_empty = rs("weight_empty")
			weight_empty_units = rs("weight_empty_units")
			height_length = rs("height_length")
			height_length_units = rs("height_length_units")
			inside_diameter = rs("inside_diameter")
			inside_diameter_units = rs("inside_diameter_units")
			shell_material = rs("shell_material")
			shell_thickness = rs("shell_thickness")
			shell_thickness_units = rs("shell_thickness_units")
			shell_min_thickness = rs("shell_min_thickness")
			shell_min_thickness_units = rs("shell_min_thickness_units")
			head_material = rs("head_material")
			head_thickness = rs("head_thickness")
			head_thickness_units = rs("head_thickness_units")
			head_min_thickness = rs("head_min_thickness")
			head_min_thickness_units = rs("head_min_thickness_units")
			lining_material = rs("lining_material")
			lining_thickness = rs("lining_thickness")
			lining_thickness_units = rs("lining_thickness_units")
			jacket_material = rs("jacket_material")
			jacket_thickness = rs("jacket_thickness")
			jacket_thickness_units = rs("jacket_thickness_units")
			mawp = rs("mawp")
			mawp_units = rs("mawp_units")
			shell_test_press = rs("shell_test_press")
			shell_test_press_units = rs("shell_test_press_units")
			jacket_test_press = rs("jacket_test_press")
			jacket_test_press_units = rs("jacket_test_press_units")
			date_built = rs("date_built")
			national_board_number = rs("national_board_number")
			lining_mfgr = rs("lining_mfgr")
			manufacturer = rs("manufacturer")
			mfgr_serial_number = rs("mfgr_serial_number")
			drawing_number = rs("drawing_number")
			jacket_type_description = rs("jacket_type_description")

		End If
		rs.Close
			
		'Determine the previous and next inspection dates.
		sqlString = "SELECT inspection_date,set_frequency,set_frequency_units " & _
				"FROM tank_inspection_data " & _
				"WHERE inspection_date=" & _
				"(SELECT MAX(inspection_date) FROM tank_inspection_data " & _
				"WHERE equipment_item_id=" & Request("itemID") & ")"
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			If Not IsNull(rs("inspection_date")) Then
				previous_inspection_date = FormatDateTime(rs("inspection_date"),2)
			Else
				previous_inspection_date = "NONE"
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
			previous_inspection_date = "NONE"
			set_frequency = 0
			set_frequency_units = ""
		End If
		rs.Close
		'If set frequency or its units are not specified, look the default values
		'up in the equipment types table.
		If set_frequency = 0 Or set_frequency_units = "" Then
			sqlString = "SELECT inspection_interval,inspection_interval_units " & _
					"FROM equipment_types WHERE equipment_type_name='Tank'"
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
	If LCase(Request("print")) = "true" Then
		Response.Write "<tr>"
		Response.Write "<td id='formtd' class='noprint' colspan='2' style='font-size:10pt;text-align:center'>"
		Response.Write "<a href='javascript: window.print();'>Print</a></td>"
		Response.Write "</tr>"
	End If
	Response.Write "<tr>"
	Response.Write "<td style='text-align:left;vertical-align:top;font-size:12pt;font-weight:bold'>Vessel - Inspection Report</td>"
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
	Response.Write "<td style='width:55%'>"
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' style='width:40%'>Test Performed</td>"
	Response.Write "<td id='grouptd' style='width:10%;padding-left:5px'>X</td>"
	Response.Write "<td id='grouptd' style='width:50%'>Type/Result</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Shell Hydro Press:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='shell_hydro_press_test_performed' name='shell_hydro_press_test_performed' value='1' /></td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Jacket Hydro Press:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='jacket_hydro_press_test_performed' name='jacket_hydro_press_test_performed' value='1' /></td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Vacuum Test:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='vacuum_test_performed' name='vacuum_test_performed' value='1' /></td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Visual (I/E/B):</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='visual_test_performed' name='visual_test_performed' value='1' /></td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Shell Ultrasonic:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='shell_ultrasonic_test_performed' name='shell_ultrasonic_test_performed' value='1' /></td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Jacket Ultrasonic:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='jacket_ultrasonic_test_performed' name='jacket_ultrasonic_test_performed' value='1' /></td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Radiographic:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='radiographic_test_performed' name='radiographic_test_performed' value='1' /></td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Magnetic Particle:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='magnetic_particle_test_performed' name='magnetic_particle_test_performed' value='1' /></td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Dye Penetrant:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='dye_penetrant_test_performed' name='dye_penetrant_test_performed' value='1' /></td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Spark Test:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='spark_test_performed' name='spark_test_performed' value='1' /></td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Other:</td>"
	Response.Write "<td id='formtd'>"
	Response.Write "<input type='checkbox' class='checkbox' id='other_test_performed' name='other_test_performed' value='1' /></td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Contractor Used:</td>"
	Response.Write "<td id='blanktd' colspan='2'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='3'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd'>Equipment Condition</td>"
	Response.Write "<td id='grouptd'>OK</td>"
	Response.Write "<td id='grouptd'>Comment</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Internal Corrosion:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>External Corrosion:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Nozzles:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Gasket Surfaces:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Weld Seams:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Lining:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Baffles & Supports:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Dip Tubes:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Agitator:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Piping & Valves:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Relief Devices:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Ladder/Handrail:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Reinforcing Rings:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Foundation - Dike:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Paint:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Insulation:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Jacket:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Nameplate Intact:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>MWO Number:</td>"
	Response.Write "<td id='blanktd' colspan='2'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "</td>"
	
	Response.Write "<td style='width:45%;vertical-align:top'>"
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' style='width:50%'>Vessel Design</td>"
	Response.Write "<td id='formtd' style='width:15%'>&nbsp;</td>"
	Response.Write "<td id='formtd' style='width:35%'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Relief Device:</td>"
	Response.Write "<td id='formtd' colspan='2'>" & relief_device & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Relief Device Pressure:</td>"
	Response.Write "<td id='formtd'>" & relief_device_pressure & "</td>"
	Response.Write "<td id='formtd'>" & relief_device_pressure_units & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Lethal Service:</td>"
	Response.Write "<td id='formtd' colspan='2'>" & lethal_service & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Capacity:</td>"
	Response.Write "<td id='formtd'>" & capacity & "</td>"
	Response.Write "<td id='formtd'>" & capacity_units & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Weight Empty:</td>"
	Response.Write "<td id='formtd'>" & weight_empty & "</td>"
	Response.Write "<td id='formtd'>" & weight_empty_units & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Height/Length:</td>"
	Response.Write "<td id='formtd'>" & height_length & "</td>"
	Response.Write "<td id='formtd'>" & height_length_units & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Inside Diameter:</td>"
	Response.Write "<td id='formtd'>" & inside_diameter & "</td>"
	Response.Write "<td id='formtd'>" & inside_diameter_units & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Shell Material:</td>"
	Response.Write "<td id='formtd' colspan='2'>" & shell_material & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Shell Thickness:</td>"
	Response.Write "<td id='formtd'>" & shell_thickness & "</td>"
	Response.Write "<td id='formtd'>" & shell_thickness_units & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Shell Min. Thick:</td>"
	Response.Write "<td id='formtd'>" & shell_min_thickness & "</td>"
	Response.Write "<td id='formtd'>" & shell_min_thickness_units & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Head Material:</td>"
	Response.Write "<td id='formtd' colspan='2'>" & head_material & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Head Thickness:</td>"
	Response.Write "<td id='formtd'>" & head_thickness & "</td>"
	Response.Write "<td id='formtd'>" & head_thickness_units & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Head Min. Thickness:</td>"
	Response.Write "<td id='formtd'>" & head_min_thickness & "</td>"
	Response.Write "<td id='formtd'>" & head_min_thickness_units & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Lining Material:</td>"
	Response.Write "<td id='formtd' colspan='2'>" & lining_material & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Lining Thickness:</td>"
	Response.Write "<td id='formtd'>" & lining_thickness & "</td>"
	Response.Write "<td id='formtd'>" & lining_thickness_units & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Jacket Material:</td>"
	Response.Write "<td id='formtd' colspan='2'>" & jacket_material & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Jacket Thickness:</td>"
	Response.Write "<td id='formtd'>" & jacket_thickness & "</td>"
	Response.Write "<td id='formtd'>" & jacket_thickness_units & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>MAWP:</td>"
	Response.Write "<td id='formtd'>" & mawp & "</td>"
	Response.Write "<td id='formtd'>" & mawp_units & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Shell Test Press:</td>"
	Response.Write "<td id='formtd'>" & shell_test_press & "</td>"
	Response.Write "<td id='formtd'>" & shell_test_press_units & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Jacket Test Press:</td>"
	Response.Write "<td id='formtd'>" & jacket_test_press & "</td>"
	Response.Write "<td id='formtd'>" & jacket_test_press_units & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Year Built:</td>"
	Response.Write "<td id='formtd' colspan='2'>" & date_built & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>National Board No:</td>"
	Response.Write "<td id='formtd' colspan='2'>" & national_board_number & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Lining MFGR:</td>"
	Response.Write "<td id='formtd' colspan='2'>" & lining_mfgr & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Manufacturer:</td>"
	Response.Write "<td id='formtd' colspan='2'>" & manufacturer & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>MFGR Serial No:</td>"
	Response.Write "<td id='formtd' colspan='2'>" & mfgr_serial_number & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Drawing No:</td>"
	Response.Write "<td id='formtd' colspan='2'>" & drawing_number & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Jacket Type Desc:</td>"
	Response.Write "<td id='formtd' colspan='2'>" & jacket_type_description & "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:45%;border:1px solid black'>"
	Response.Write "<table style='width:100%'>"
	Response.Write "<tr>"
	Response.Write "<td id='smalltd' style='width:40%'>Sketches Attached:</td>"
	Response.Write "<td id='blanktd' style='width:10%'>&nbsp;</td>"
	Response.Write "<td id='smalltd' style='width:40%'>Repair Required:</td>"
	Response.Write "<td id='blanktd' style='width:10%'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='smalltd'>NDT Reports Attached:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='smalltd'>Type:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='smalltd'>Policy Verify:</td>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='smalltd' colspan='2'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "</td>"
	Response.Write "<td id='formtd' style='width:55%'>"
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='smalltd' style='width:31%'>Inspection Company:</td>"
	Response.Write "<td id='blanktd' colspan='4'>&nbsp;</td>"
	Response.Write "<td id='formtd' style='width:5%'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='smalltd' style='width:31%'>Inspector:</td>"
	Response.Write "<td id='blanktd' colspan='4'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='smalltd'>Next Inspection Date:</td>"
	Response.Write "<td id='smalltd' colspan='5'>"
	Response.Write next_inspection_due & "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='smalltd'>Previous Inspection:</td>"
	Response.Write "<td id='smalltd' style='width:21%'>" & previous_inspection_date & "</td>"
	Response.Write "<td id='smalltd' style='width:23%'>Set Frequency:</td>"
	Response.Write "<td id='smalltd' style='width:11%'>"
	Response.Write set_frequency & "</td>"
	Response.Write "<td id='smalltd' style='width:8%'>"
	Response.Write set_frequency_units & "</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "</td></tr>"
	Response.Write "</table>"

	Response.Write "<br />"
	Response.Write "<div style='page-break-before:always;font-size:1;margin:0;border:0'><span style='visibility:hidden'>-</span></div>"
	
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' style='width:95%'>SPECIFIC TEST DATA & FINDINGS:</td>"
	Response.Write "<td id='formtd' style='width:5%'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='2'>SUMMARY & RECOMMENDATIONS:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='2'>COMMENTS:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='2'>DISCREPANCY:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2' style='font-weight:bold'>&nbsp;&nbsp;COMMENTS:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2' style='font-weight:bold'>&nbsp;&nbsp;FOLLOW-UP:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='blanktd'>&nbsp;</td>"
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	
	Response.Write "<div style='text-align:left'>"
	Response.Write "<table style='width:50%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2' style='font-weight:bold'>Reference:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:55%;font-weight:bold'>Location</td>"
	Response.Write "<td id='formtd' style='width:45%;font-weight:bold'>Reading</td>"
	Response.Write "</tr>"
	Dim count
	For count = 1 To 10
		Response.Write "<tr>"
		Response.Write "<td id='blanktd'>&nbsp;</td>"
		Response.Write "<td id='blanktd'>&nbsp;</td>"
		Response.Write "</tr>"
	Next
	
	Response.Write "</table>"
	Response.Write "</div>"
		
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
