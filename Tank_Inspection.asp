<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<html>
<head>
<script language="javascript">
var needToConfirm = false;

function openhelp() {
 window.open("Equipment Inspection Database Users Guide.doc","userguide");
}
function addDate(d,interval,unit) {
 //If the previous inspection date does not exist, use the current date.
 if (d != '') {
  t = new Date(d);
 } else {
  t = new Date();
 }
 //Make javascript treat the interval as a number.
 interval = interval*1;
 //Add the appropriate interval to the date.
 if (unit=="days") {
  t.setDate(t.getDate()+interval);
 } else if (unit=="months") {
  t.setMonth(t.getMonth()+interval);
 } else if (unit=="years") {
  t.setFullYear(t.getFullYear()+interval);
 }
 //Put the result in the next inspection field.
 document.form1.next_inspection_due.value=t.getMonth()+1+"/"+t.getDate()+"/"+t.getFullYear();
 needToConfirm=true;
}
function isNumeric(val) {
 var checkOk = "0123456789,.-";
 var valOk = true;
 for (i=0;i<val.length;i++) {
  ch = val.charAt(i);
  for (j=0;j<checkOk.length;j++) {
   if (ch==checkOk.charAt(j)) {
    break;
   }
   if (j==checkOk.length-1) {
    valOk = false;
    break;
   }
  }
 }
 return valOk;
}
function chkNumeric(id) {
 var valOk = true;
 var val = id.value;
 valOk = isNumeric(val);
 if (valOk==false) {
  id.style.color="red";
  alert('Invalid number entered');
 } else {
  id.style.color="black";
 }
}
function isDate(val) {
 var checkOk = "0123456789/-";
 var valOk = true;
 for (i=0;i<val.length;i++) {
  ch = val.charAt(i);
  for (j=0;j<checkOk.length;j++) {
   if (ch==checkOk.charAt(j)) {
    break;
   }
   if (j==checkOk.length-1) {
    valOk = false;
    break;
   }
  }
 }
 if (val.length<6) {
  valOk = false;
 }
 return valOk;
}
function chkDate(id) {
 var valOk = true;
 var val = id.value;
 valOk = isDate(val);
 if (valOk==false) {
  id.style.color="red";
  alert('Invalid date entered');
 } else {
  id.style.color="black";
 }
} 
function opentechdata() {
 window.open("http://mogsb8/inspections/tank_technicaldata.asp?itemID="+document.form1.itemID.value+"&edit=false","TechnicalData");
}
function saveData() {
 needToConfirm=false;
 document.form1.submit();
}
function setupdate() {
 needToConfirm=true;
}
function editReading(id) {
 document.form1.action="inspectionaction.asp?action=editReading&readingID="+id;
 document.form1.submit();
}
function addReading() {
 document.form1.action="inspectionaction.asp?action=addReading";
 document.form1.submit();
}
function updateReading(id) {
 document.form1.action="inspectionaction.asp?action=updateReading&readingID="+id;
 document.form1.submit();
}
function warn() {
 if (needToConfirm==true) {
  return "You have changed the data on this form and not submitted it.";
 }
}
window.onbeforeunload = warn;
<!--#include file="datepicker.js"-->
</script>
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
' Keith Brooks - Tuesday, March 1, 2011
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
Dim field_disabled
'Read-only items
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
'Inspection items
Dim inspection_date
Dim shell_hydro_press_test_performed
Dim shell_hydro_press_test_type_result
Dim jacket_hydro_press_test_performed
Dim jacket_hydro_press_test_type_result
Dim vacuum_test_performed
Dim vacuum_test_type_result
Dim visual_test_performed
Dim visual_test_type_result
Dim shell_ultrasonic_test_performed
Dim shell_ultrasonic_test_type_result
Dim jacket_ultrasonic_test_performed
Dim jacket_ultrasonic_test_type_result
Dim radiographic_test_performed
Dim radiographic_test_type_result
Dim magnetic_particle_test_performed
Dim magnetic_particle_test_type_result
Dim dye_penetrant_test_performed
Dim dye_penetrant_test_type_result
Dim spark_test_performed
Dim spark_test_type_result
Dim other_test_performed
Dim other_test_type_result
Dim contractor_used
Dim internal_corrosion_ok
Dim internal_corrosion_comment
Dim external_corrosion_ok
Dim external_corrosion_comment
Dim nozzles_ok
Dim nozzles_comment
Dim gasket_surfaces_ok
Dim gasket_surfaces_comment
Dim weld_seams_ok
Dim weld_seams_comment
Dim lining_ok
Dim lining_comment
Dim baffles_supports_ok
Dim baffles_supports_comment
Dim dip_tubes_ok
Dim dip_tubes_comment
Dim agitator_ok
Dim agitator_comment
Dim piping_valves_ok
Dim piping_valves_comment
Dim relief_devices_ok
Dim relief_devices_comment
Dim ladder_handrail_ok
Dim ladder_handrail_comment
Dim reinforcing_rings_ok
Dim reinforcing_rings_comment
Dim foundation_dike_ok
Dim foundation_dike_comment
Dim paint_ok
Dim paint_comment
Dim insulation_ok
Dim insulation_comment
Dim jacket_ok
Dim jacket_comment
Dim nameplate_intact_ok
Dim nameplate_intact_comment
Dim mwo_number
Dim specific_test_data_findings
Dim summary_recommendations
Dim comments
Dim discrepency_comments
Dim discrepency_followup
Dim sketches_attached
Dim ndt_reports_attached
Dim policy_verify
Dim repair_required
Dim repair_type
Dim inspection_company
Dim inspected_by
Dim next_inspection_due
Dim previous_inspection_date
Dim set_frequency
Dim set_frequency_units

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("inspections", "tank_inspection", currentuser)
If access <> "none" Then

	If LCase(Request("edit")) = "true" And (access = "write" Or access = "delete") Then
		editMode = True
		field_disabled = ""
		If session("focus") = "" Then
			Response.Write "<body  onload='document.form1.inspection_date.focus();'>"
		Else
			Response.Write "<body  onload='document.form1." & session("focus") & ".focus();'>"
		End If
	Else
		editMode = False
		field_disabled = "disabled"
		If LCase(Request("print")) = "true" Then
'			Response.Write "<body style='background-color:white' onload='window.print();window.close();'>"
			Response.Write "<body style='background-color:white'>"
		Else
			Response.Write "<body>"
		End If
	End If
		
	'Define the ado connection and recordset objects.
	set cn = CreateObject("adodb.connection")
	cn.Open = DBString
	set rs = CreateObject("adodb.recordset")

	Response.Write "<form id='form1' name='form1' action='inspectionaction.asp' method='post'>"
	
	'Save the equipment type for use by the action page.
	Response.Write "<input type='hidden' id='equipType' name='equipType' value='tank' />"
	
	'If this is an existing inspection, get the record from the database.
	If Request("inspectionID") <> "" Then
		'Save the value of the inspection id.
		Response.Write "<input type='hidden' id='inspectionID' name='inspectionID' value='" & Request("inspectionID") & "' />"

		'Get the existing inspection data.
		sqlString = "SELECT * FROM tank_inspection_data " & _
				"WHERE inspection_data_id=" & Request("inspectionID")
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			itemID = rs("equipment_item_id")

			'Save the value of the item id.
			Response.Write "<input type='hidden' id='itemID' name='itemID' value='" & itemID & "' />"

			'Get the static technical data for this item from the database.
			sqlString = "SELECT equipment_item_tag,equipment_item_name,assembly,area,b.* " & _
					"FROM equipment_items a INNER JOIN tank_technical_data b " & _
					"ON a.equipment_item_id=b.equipment_item_id " & _
					"WHERE a.equipment_item_id=" & itemID
			Set rs2 = CreateObject("adodb.recordset")
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				'Fill in the static technical data variables.
				rs2.MoveFirst
				equipment_item_tag = rs2("equipment_item_tag")
				equipment_item_name = rs2("equipment_item_name")
				area = rs2("area")
				assembly = rs2("assembly")
				state_number = rs2("state_number")
				relief_device = rs2("relief_device")
				relief_device_pressure = rs2("relief_device_pressure")
				relief_device_pressure_units = rs2("relief_device_pressure_units")
				lethal_service = rs2("lethal_service")
				capacity = rs2("capacity")
				capacity_units = rs2("capacity_units")
				weight_empty = rs2("weight_empty")
				weight_empty_units = rs2("weight_empty_units")
				height_length = rs2("height_length")
				height_length_units = rs2("height_length_units")
				inside_diameter = rs2("inside_diameter")
				inside_diameter_units = rs2("inside_diameter_units")
				shell_material = rs2("shell_material")
				shell_thickness = rs2("shell_thickness")
				shell_thickness_units = rs2("shell_thickness_units")
				shell_min_thickness = rs2("shell_min_thickness")
				shell_min_thickness_units = rs2("shell_min_thickness_units")
				head_material = rs2("head_material")
				head_thickness = rs2("head_thickness")
				head_thickness_units = rs2("head_thickness_units")
				head_min_thickness = rs2("head_min_thickness")
				head_min_thickness_units = rs2("head_min_thickness_units")
				lining_material = rs2("lining_material")
				lining_thickness = rs2("lining_thickness")
				lining_thickness_units = rs2("lining_thickness_units")
				jacket_material = rs2("jacket_material")
				jacket_thickness = rs2("jacket_thickness")
				jacket_thickness_units = rs2("jacket_thickness_units")
				mawp = rs2("mawp")
				mawp_units = rs2("mawp_units")
				shell_test_press = rs2("shell_test_press")
				shell_test_press_units = rs2("shell_test_press_units")
				jacket_test_press = rs2("jacket_test_press")
				jacket_test_press_units = rs2("jacket_test_press_units")
				date_built = rs2("date_built")
				national_board_number = rs2("national_board_number")
				lining_mfgr = rs2("lining_mfgr")
				manufacturer = rs2("manufacturer")
				mfgr_serial_number = rs2("mfgr_serial_number")
				drawing_number = rs2("drawing_number")
				jacket_type_description = rs2("jacket_type_description")

			End If
			rs2.Close
			Set rs2 = Nothing
			
			'Fill in the existing inspection data variables.
			If IsNull(rs("inspection_date")) Then
				inspection_date = ""
			Else
				inspection_date = FormatDateTime(rs("inspection_date"),2)
			End If
			shell_hydro_press_test_performed = rs("shell_hydro_press_test_performed")
			shell_hydro_press_test_type_result = rs("shell_hydro_press_test_type_result")
			jacket_hydro_press_test_performed = rs("jacket_hydro_press_test_performed")
			jacket_hydro_press_test_type_result = rs("jacket_hydro_press_test_type_result")
			vacuum_test_performed = rs("vacuum_test_performed")
			vacuum_test_type_result = rs("vacuum_test_type_result")
			visual_test_performed = rs("visual_test_performed")
			visual_test_type_result = rs("visual_test_type_result")
			shell_ultrasonic_test_performed = rs("shell_ultrasonic_test_performed")
			shell_ultrasonic_test_type_result = rs("shell_ultrasonic_test_type_result")
			jacket_ultrasonic_test_performed = rs("jacket_ultrasonic_test_performed")
			jacket_ultrasonic_test_type_result = rs("jacket_ultrasonic_test_type_result")
			radiographic_test_performed = rs("radiographic_test_performed")
			radiographic_test_type_result = rs("radiographic_test_type_result")
			magnetic_particle_test_performed = rs("magnetic_particle_test_performed")
			magnetic_particle_test_type_result = rs("magnetic_particle_test_type_result")
			dye_penetrant_test_performed = rs("dye_penetrant_test_performed")
			dye_penetrant_test_type_result = rs("dye_penetrant_test_type_result")
			spark_test_performed = rs("spark_test_performed")
			spark_test_type_result = rs("spark_test_type_result")
			other_test_performed = rs("other_test_performed")
			other_test_type_result = rs("other_test_type_result")
			contractor_used = rs("contractor_used")
			internal_corrosion_ok = rs("internal_corrosion_ok")
			internal_corrosion_comment = rs("internal_corrosion_comment")
			external_corrosion_ok = rs("external_corrosion_ok")
			external_corrosion_comment = rs("external_corrosion_comment")
			nozzles_ok = rs("nozzles_ok")
			nozzles_comment = rs("nozzles_comment")
			gasket_surfaces_ok = rs("gasket_surfaces_ok")
			gasket_surfaces_comment = rs("gasket_surfaces_comment")
			weld_seams_ok = rs("weld_seams_ok")
			weld_seams_comment = rs("weld_seams_comment")
			lining_ok = rs("lining_ok")
			lining_comment = rs("lining_comment")
			baffles_supports_ok = rs("baffles_supports_ok")
			baffles_supports_comment = rs("baffles_supports_comment")
			dip_tubes_ok = rs("dip_tubes_ok")
			dip_tubes_comment = rs("dip_tubes_comment")
			agitator_ok = rs("agitator_ok")
			agitator_comment = rs("agitator_comment")
			piping_valves_ok = rs("piping_valves_ok")
			piping_valves_comment = rs("piping_valves_comment")
			relief_devices_ok = rs("relief_devices_ok")
			relief_devices_comment = rs("relief_devices_comment")
			ladder_handrail_ok = rs("ladder_handrail_ok")
			ladder_handrail_comment = rs("ladder_handrail_comment")
			reinforcing_rings_ok = rs("reinforcing_rings_ok")
			reinforcing_rings_comment = rs("reinforcing_rings_comment")
			foundation_dike_ok = rs("foundation_dike_ok")
			foundation_dike_comment = rs("foundation_dike_comment")
			paint_ok = rs("paint_ok")
			paint_comment = rs("paint_comment")
			insulation_ok = rs("insulation_ok")
			insulation_comment = rs("insulation_comment")
			jacket_ok = rs("jacket_ok")
			jacket_comment = rs("jacket_comment")
			nameplate_intact_ok = rs("nameplate_intact_ok")
			nameplate_intact_comment = rs("nameplate_intact_comment")
			mwo_number = rs("mwo_number")
			specific_test_data_findings = rs("specific_test_data_findings")
			summary_recommendations = rs("summary_recommendations")
			comments = rs("comments")
			discrepency_comments = rs("discrepency_comments")
			discrepency_followup = rs("discrepency_followup")
			sketches_attached = rs("sketches_attached")
			ndt_reports_attached = rs("ndt_reports_attached")
			policy_verify = rs("policy_verify")
			repair_required = rs("repair_required")
			repair_type = rs("repair_type")
			inspection_company = rs("inspection_company")
			inspected_by = rs("inspected_by")
			If IsNull(rs("next_inspection_due")) Then
				next_inspection_due = ""
			Else
				next_inspection_due = FormatDateTime(rs("next_inspection_due"),2)
			End If
			If IsNull(rs("previous_inspection_date")) Then
				previous_inspection_date = "NONE"
			Else
				previous_inspection_date = FormatDateTime(rs("previous_inspection_date"),2)
			End If
			set_frequency = rs("set_frequency")
			set_frequency_units = rs("set_frequency_units")
			
		End If
	Else
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
			
			'Initialize the inspection data variables.
			inspection_date = Date
			shell_hydro_press_test_performed = 0
			shell_hydro_press_test_type_result = ""
			jacket_hydro_press_test_performed = 0
			jacket_hydro_press_test_type_result = ""
			vacuum_test_performed = 0
			vacuum_test_type_result = ""
			visual_test_performed = 0
			visual_test_type_result = ""
			shell_ultrasonic_test_performed = 0
			shell_ultrasonic_test_type_result = ""
			jacket_ultrasonic_test_performed = 0
			jacket_ultrasonic_test_type_result = ""
			radiographic_test_performed = 0
			radiographic_test_type_result = ""
			magnetic_particle_test_performed = 0
			magnetic_particle_test_type_result = ""
			dye_penetrant_test_performed = 0
			dye_penetrant_test_type_result = ""
			spark_test_performed = 0
			spark_test_type_result = ""
			other_test_performed = 0
			other_test_type_result = ""
			contractor_used = ""
			internal_corrosion_ok = ""
			internal_corrosion_comment = ""
			external_corrosion_ok = ""
			external_corrosion_comment = ""
			nozzles_ok = ""
			nozzles_comment = ""
			gasket_surfaces_ok = ""
			gasket_surfaces_comment = ""
			weld_seams_ok = ""
			weld_seams_comment = ""
			lining_ok = ""
			lining_comment = ""
			baffles_supports_ok = ""
			baffles_supports_comment = ""
			dip_tubes_ok = ""
			dip_tubes_comment = ""
			agitator_ok = ""
			agitator_comment = ""
			piping_valves_ok = ""
			piping_valves_comment = ""
			relief_devices_ok = ""
			relief_devices_comment = ""
			ladder_handrail_ok = ""
			ladder_handrail_comment = ""
			reinforcing_rings_ok = ""
			reinforcing_rings_comment = ""
			foundation_dike_ok = ""
			foundation_dike_comment = ""
			paint_ok = ""
			paint_comment = ""
			insulation_ok = ""
			insulation_comment = ""
			jacket_ok = ""
			jacket_comment = ""
			nameplate_intact_ok = ""
			nameplate_intact_comment = ""
			mwo_number = ""
			specific_test_data_findings = ""
			summary_recommendations = ""
			comments = ""
			discrepency_comments = ""
			discrepency_followup = ""
			sketches_attached = ""
			ndt_reports_attached = ""
			policy_verify = ""
			repair_required = ""
			repair_type = ""
			inspection_company = ""
			inspected_by = ""
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
	End If
	
	'Save the previous inspection date in a hidden tag so it can be saved with
	'the rest of the data.
	Response.Write "<input type='hidden' id='previous_inspection_date' name='previous_inspection_date' value='" & previous_inspection_date & "' />"

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
	Else
		Response.Write "<tr>"
		Response.Write "<td class='noprint' style='text-align:left;vertical-align:top;width:50%'><a href='default.asp'>Home</a></td>"
		Response.Write "<td class='noprint' style='text-align:right;vertical-align:top;width:50%'><a href='' onclick='openhelp();return false;' title='Open the User Guide'>Help</a></td>"
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
	Response.Write "<td id='formtd' style='width:25%'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' id='inspection_date' name='inspection_date' value='" & inspection_date & "' onchange='chkDate(this);setupdate();' />"
		Response.Write "<a href='javascript: void(0);' onclick='displayDatePicker(""inspection_date"");setupdate();return false;'><img src='../images/calendar.bmp' alt='Calendar' style='vertical-align:top'></a></td>"
	Else
		Response.Write inspection_date & "</td>"
	End If
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
	Response.Write "<tr><td width:'100%'>&nbsp;"
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
	If shell_hydro_press_test_performed = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='shell_hydro_press_test_performed' name='shell_hydro_press_test_performed' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='shell_hydro_press_test_performed' name='shell_hydro_press_test_performed' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='shell_hydro_press_test_type_result' name='shell_hydro_press_test_type_result' value='" & shell_hydro_press_test_type_result & "' onchange='setupdate();' /></td>"
	Else
		Response.Write shell_hydro_press_test_type_result & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Jacket Hydro Press:</td>"
	Response.Write "<td id='formtd'>"
	If jacket_hydro_press_test_performed = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='jacket_hydro_press_test_performed' name='jacket_hydro_press_test_performed' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='jacket_hydro_press_test_performed' name='jacket_hydro_press_test_performed' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='jacket_hydro_press_test_type_result' name='jacket_hydro_press_test_type_result' value='" & jacket_hydro_press_test_type_result & "' onchange='setupdate();' /></td>"
	Else
		Response.Write jacket_hydro_press_test_type_result & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Vacuum Test:</td>"
	Response.Write "<td id='formtd'>"
	If vacuum_test_performed = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='vacuum_test_performed' name='vacuum_test_performed' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='vacuum_test_performed' name='vacuum_test_performed' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='vacuum_test_type_result' name='vacuum_test_type_result' value='" & vacuum_test_type_result & "' onchange='setupdate();' /></td>"
	Else
		Response.Write vacuum_test_type_result & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Visual (I/E/B):</td>"
	Response.Write "<td id='formtd'>"
	If visual_test_performed = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='visual_test_performed' name='visual_test_performed' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='visual_test_performed' name='visual_test_performed' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='visual_test_type_result' name='visual_test_type_result' value='" & visual_test_type_result & "' onchange='setupdate();' /></td>"
	Else
		Response.Write visual_test_type_result & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Shell Ultrasonic:</td>"
	Response.Write "<td id='formtd'>"
	If shell_ultrasonic_test_performed = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='shell_ultrasonic_test_performed' name='shell_ultrasonic_test_performed' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='shell_ultrasonic_test_performed' name='shell_ultrasonic_test_performed' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='shell_ultrasonic_test_type_result' name='shell_ultrasonic_test_type_result' value='" & shell_ultrasonic_test_type_result & "' onchange='setupdate();' /></td>"
	Else
		Response.Write shell_ultrasonic_test_type_result & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Jacket Ultrasonic:</td>"
	Response.Write "<td id='formtd'>"
	If jacket_ultrasonic_test_performed = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='jacket_ultrasonic_test_performed' name='jacket_ultrasonic_test_performed' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='jacket_ultrasonic_test_performed' name='jacket_ultrasonic_test_performed' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='jacket_ultrasonic_test_type_result' name='jacket_ultrasonic_test_type_result' value='" & jacket_ultrasonic_test_type_result & "' onchange='setupdate();' /></td>"
	Else
		Response.Write jacket_ultrasonic_test_type_result & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Radiographic:</td>"
	Response.Write "<td id='formtd'>"
	If radiographic_test_performed = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='radiographic_test_performed' name='radiographic_test_performed' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='radiographic_test_performed' name='radiographic_test_performed' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='radiographic_test_type_result' name='radiographic_test_type_result' value='" & radiographic_test_type_result & "' onchange='setupdate();' /></td>"
	Else
		Response.Write radiographic_test_type_result & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Magnetic Particle:</td>"
	Response.Write "<td id='formtd'>"
	If magnetic_particle_test_performed = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='magnetic_particle_test_performed' name='magnetic_particle_test_performed' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='magnetic_particle_test_performed' name='magnetic_particle_test_performed' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='magnetic_particle_test_type_result' name='magnetic_particle_test_type_result' value='" & magnetic_particle_test_type_result & "' onchange='setupdate();' /></td>"
	Else
		Response.Write magnetic_particle_test_type_result & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Dye Penetrant:</td>"
	Response.Write "<td id='formtd'>"
	If dye_penetrant_test_performed = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='dye_penetrant_test_performed' name='dye_penetrant_test_performed' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='dye_penetrant_test_performed' name='dye_penetrant_test_performed' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='dye_penetrant_test_type_result' name='dye_penetrant_test_type_result' value='" & dye_penetrant_test_type_result & "' onchange='setupdate();' /></td>"
	Else
		Response.Write dye_penetrant_test_type_result & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Spark Test:</td>"
	Response.Write "<td id='formtd'>"
	If spark_test_performed = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='spark_test_performed' name='spark_test_performed' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='spark_test_performed' name='spark_test_performed' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='spark_test_type_result' name='spark_test_type_result' value='" & spark_test_type_result & "' onchange='setupdate();' /></td>"
	Else
		Response.Write spark_test_type_result & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Other:</td>"
	Response.Write "<td id='formtd'>"
	If other_test_performed = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='other_test_performed' name='other_test_performed' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='other_test_performed' name='other_test_performed' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='other_test_type_result' name='other_test_type_result' value='" & other_test_type_result & "' onchange='setupdate();' /></td>"
	Else
		Response.Write other_test_type_result & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Contractor Used:</td>"
	Response.Write "<td id='formtd' colspan='2'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:91%' id='contractor_used' name='contractor_used' value='" & contractor_used & "' onchange='setupdate();' /></td>"
	Else
		Response.Write contractor_used & "</td>"
	End If
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
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<select id='internal_corrosion_ok' name='internal_corrosion_ok' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If internal_corrosion_ok = "Yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If internal_corrosion_ok = "No" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		If internal_corrosion_ok = "NA" Then
			Response.Write "<option value='NA' selected>NA"
		Else
			Response.Write "<option value='NA'>NA"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write internal_corrosion_ok & "</td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='internal_corrosion_comment' name='internal_corrosion_comment' value='" & internal_corrosion_comment & "' onchange='setupdate();' /></td>"
	Else
		Response.Write internal_corrosion_comment & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>External Corrosion:</td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<select id='external_corrosion_ok' name='external_corrosion_ok' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If external_corrosion_ok = "Yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If external_corrosion_ok = "No" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		If external_corrosion_ok = "NA" Then
			Response.Write "<option value='NA' selected>NA"
		Else
			Response.Write "<option value='NA'>NA"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write external_corrosion_ok & "</td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='external_corrosion_comment' name='external_corrosion_comment' value='" & external_corrosion_comment & "' onchange='setupdate();' /></td>"
	Else
		Response.Write external_corrosion_comment & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Nozzles:</td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<select id='nozzles_ok' name='nozzles_ok' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If nozzles_ok = "Yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If nozzles_ok = "No" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		If nozzles_ok = "NA" Then
			Response.Write "<option value='NA' selected>NA"
		Else
			Response.Write "<option value='NA'>NA"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write nozzles_ok & "</td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='nozzles_comment' name='nozzles_comment' value='" & nozzles_comment & "' onchange='setupdate();' /></td>"
	Else
		Response.Write nozzles_comment & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Gasket Surfaces:</td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<select id='gasket_surfaces_ok' name='gasket_surfaces_ok' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If gasket_surfaces_ok = "Yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If gasket_surfaces_ok = "No" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		If gasket_surfaces_ok = "NA" Then
			Response.Write "<option value='NA' selected>NA"
		Else
			Response.Write "<option value='NA'>NA"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write gasket_surfaces_ok & "</td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='gasket_surfaces_comment' name='gasket_surfaces_comment' value='" & gasket_surfaces_comment & "' onchange='setupdate();' /></td>"
	Else
		Response.Write gasket_surfaces_comment & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Weld Seams:</td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<select id='weld_seams_ok' name='weld_seams_ok' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If weld_seams_ok = "Yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If weld_seams_ok = "No" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		If weld_seams_ok = "NA" Then
			Response.Write "<option value='NA' selected>NA"
		Else
			Response.Write "<option value='NA'>NA"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write weld_seams_ok & "</td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='weld_seams_comment' name='weld_seams_comment' value='" & weld_seams_comment & "' onchange='setupdate();' /></td>"
	Else
		Response.Write weld_seams_comment & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Lining:</td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<select id='lining_ok' name='lining_ok' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If lining_ok = "Yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If lining_ok = "No" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		If lining_ok = "NA" Then
			Response.Write "<option value='NA' selected>NA"
		Else
			Response.Write "<option value='NA'>NA"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write lining_ok & "</td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='lining_comment' name='lining_comment' value='" & lining_comment & "' onchange='setupdate();' /></td>"
	Else
		Response.Write lining_comment & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Baffles & Supports:</td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<select id='baffles_supports_ok' name='baffles_supports_ok' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If baffles_supports_ok = "Yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If baffles_supports_ok = "No" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		If baffles_supports_ok = "NA" Then
			Response.Write "<option value='NA' selected>NA"
		Else
			Response.Write "<option value='NA'>NA"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write baffles_supports_ok & "</td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='baffles_supports_comment' name='baffles_supports_comment' value='" & baffles_supports_comment & "' onchange='setupdate();' /></td>"
	Else
		Response.Write baffles_supports_comment & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Dip Tubes:</td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<select id='dip_tubes_ok' name='dip_tubes_ok' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If dip_tubes_ok = "Yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If dip_tubes_ok = "No" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		If dip_tubes_ok = "NA" Then
			Response.Write "<option value='NA' selected>NA"
		Else
			Response.Write "<option value='NA'>NA"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write lining_ok & "</td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='dip_tubes_comment' name='dip_tubes_comment' value='" & dip_tubes_comment & "' onchange='setupdate();' /></td>"
	Else
		Response.Write dip_tubes_comment & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Agitator:</td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<select id='agitator_ok' name='agitator_ok' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If agitator_ok = "Yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If agitator_ok = "No" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		If agitator_ok = "NA" Then
			Response.Write "<option value='NA' selected>NA"
		Else
			Response.Write "<option value='NA'>NA"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write agitator_ok & "</td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='agitator_comment' name='agitator_comment' value='" & agitator_comment & "' onchange='setupdate();' /></td>"
	Else
		Response.Write agitator_comment & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Piping & Valves:</td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<select id='piping_valves_ok' name='piping_valves_ok' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If piping_valves_ok = "Yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If piping_valves_ok = "No" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		If piping_valves_ok = "NA" Then
			Response.Write "<option value='NA' selected>NA"
		Else
			Response.Write "<option value='NA'>NA"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write piping_valves_ok & "</td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='piping_valves_comment' name='piping_valves_comment' value='" & piping_valves_comment & "' onchange='setupdate();' /></td>"
	Else
		Response.Write piping_valves_comment & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Relief Devices:</td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<select id='relief_devices_ok' name='relief_devices_ok' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If relief_devices_ok = "Yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If relief_devices_ok = "No" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		If relief_devices_ok = "NA" Then
			Response.Write "<option value='NA' selected>NA"
		Else
			Response.Write "<option value='NA'>NA"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write relief_devices_ok & "</td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='relief_devices_comment' name='relief_devices_comment' value='" & relief_devices_comment & "' onchange='setupdate();' /></td>"
	Else
		Response.Write relief_devices_comment & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Ladder/Handrail:</td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<select id='ladder_handrail_ok' name='ladder_handrail_ok' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If ladder_handrail_ok = "Yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If ladder_handrail_ok = "No" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		If ladder_handrail_ok = "NA" Then
			Response.Write "<option value='NA' selected>NA"
		Else
			Response.Write "<option value='NA'>NA"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write ladder_handrail_ok & "</td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='ladder_handrail_comment' name='ladder_handrail_comment' value='" & ladder_handrail_comment & "' onchange='setupdate();' /></td>"
	Else
		Response.Write ladder_handrail_comment & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Reinforcing Rings:</td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<select id='reinforcing_rings_ok' name='reinforcing_rings_ok' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If reinforcing_rings_ok = "Yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If reinforcing_rings_ok = "No" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		If reinforcing_rings_ok = "NA" Then
			Response.Write "<option value='NA' selected>NA"
		Else
			Response.Write "<option value='NA'>NA"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write reinforcing_rings_ok & "</td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='reinforcing_rings_comment' name='reinforcing_rings_comment' value='" & reinforcing_rings_comment & "' onchange='setupdate();' /></td>"
	Else
		Response.Write reinforcing_rings_comment & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Foundation - Dike:</td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<select id='foundation_dike_ok' name='foundation_dike_ok' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If foundation_dike_ok = "Yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If foundation_dike_ok = "No" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		If foundation_dike_ok = "NA" Then
			Response.Write "<option value='NA' selected>NA"
		Else
			Response.Write "<option value='NA'>NA"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write foundation_dike_ok & "</td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='foundation_dike_comment' name='foundation_dike_comment' value='" & foundation_dike_comment & "' onchange='setupdate();' /></td>"
	Else
		Response.Write foundation_dike_comment & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Paint:</td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<select id='paint_ok' name='paint_ok' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If paint_ok = "Yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If paint_ok = "No" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		If paint_ok = "NA" Then
			Response.Write "<option value='NA' selected>NA"
		Else
			Response.Write "<option value='NA'>NA"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write paint_ok & "</td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='paint_comment' name='paint_comment' value='" & paint_comment & "' onchange='setupdate();' /></td>"
	Else
		Response.Write paint_comment & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Insulation:</td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<select id='insulation_ok' name='insulation_ok' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If insulation_ok = "Yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If insulation_ok = "No" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		If insulation_ok = "NA" Then
			Response.Write "<option value='NA' selected>NA"
		Else
			Response.Write "<option value='NA'>NA"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write insulation_ok & "</td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='insulation_comment' name='insulation_comment' value='" & insulation_comment & "' onchange='setupdate();' /></td>"
	Else
		Response.Write insulation_comment & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Jacket:</td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<select id='jacket_ok' name='jacket_ok' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If jacket_ok = "Yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If jacket_ok = "No" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		If jacket_ok = "NA" Then
			Response.Write "<option value='NA' selected>NA"
		Else
			Response.Write "<option value='NA'>NA"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write jacket_ok & "</td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='jacket_comment' name='jacket_comment' value='" & jacket_comment & "' onchange='setupdate();' /></td>"
	Else
		Response.Write jacket_comment & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Nameplate Intact:</td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<select id='nameplate_intact_ok' name='nameplate_intact_ok' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If nameplate_intact_ok = "Yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If nameplate_intact_ok = "No" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		If nameplate_intact_ok = "NA" Then
			Response.Write "<option value='NA' selected>NA"
		Else
			Response.Write "<option value='NA'>NA"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write nameplate_intact_ok & "</td>"
	End If
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='nameplate_intact_comment' name='nameplate_intact_comment' value='" & nameplate_intact_comment & "' onchange='setupdate();' /></td>"
	Else
		Response.Write nameplate_intact_comment & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>MWO Number:</td>"
	Response.Write "<td id='formtd' colspan='2'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:91%' id='mwo_number' name='mwo_number' value='" & mwo_number & "' onchange='setupdate();' /></td>"
	Else
		Response.Write mwo_number & "</td>"
	End If
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
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='2'>SPECIFIC TEST DATA & FINDINGS:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2'>"
	If editMode = True Then
		Response.Write "<textarea style='width:100%;height:40px' id='specific_test_data_findings' name='specific_test_data_findings' rows='2' cols='80' onchange='setupdate();'>" & specific_test_data_findings & "</textarea></td>"
	Else
		Response.Write specific_test_data_findings & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='2'>SUMMARY & RECOMMENDATIONS:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2'>"
	If editMode = True Then
		Response.Write "<textarea style='width:100%;height:40px' id='summary_recommendations' name='summary_recommendations' rows='2' cols='80' onchange='setupdate();'>" & summary_recommendations & "</textarea></td>"
	Else
		Response.Write summary_recommendations & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "</td></tr>"
	Response.Write "<tr><td>"
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:45%;border:1px solid black'>"
	Response.Write "<table style='width:100%'>"
	Response.Write "<tr>"
	If editMode = True Then
		Response.Write "<td id='formtd' style='width:40%'>Sketches Attached:</td>"
		Response.Write "<td id='formtd' style='width:10%'>"
		Response.Write "<select id='sketches_attached' name='sketches_attached' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If LCase(sketches_attached) = "yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If LCase(sketches_attached) = "no" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write "<td id='smalltd' style='width:40%'>Sketches Attached:</td>"
		Response.Write "<td id='smalltd' style='width:10%'>" & sketches_attached & "</td>"
	End If
	If editMode = True Then
		Response.Write "<td id='formtd' style='width:40%'>Repair Required:</td>"
		Response.Write "<td id='formtd' style='width:10%'>"
		Response.Write "<select id='repair_required' name='repair_required' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If LCase(repair_required) = "yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If LCase(repair_required) = "no" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write "<td id='smalltd' style='width:40%'>Repair Required:</td>"
		Response.Write "<td id='smalltd' style='width:10%'>"
		Response.Write repair_required & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	If editMode = True Then
		Response.Write "<td id='formtd'>NDT Reports Attached:</td>"
		Response.Write "<td id='formtd'>"
		Response.Write "<select id='ndt_reports_attached' name='ndt_reports_attached' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If LCase(ndt_reports_attached) = "yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If LCase(ndt_reports_attached) = "no" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write "<td id='smalltd'>NDT Reports Attached:</td>"
		Response.Write "<td id='smalltd'>"
		Response.Write ndt_reports_attached & "</td>"
	End If
	If editMode = True Then
		Response.Write "<td id='formtd'>Type:</td>"
		Response.Write "<td id='formtd'>"
		Response.Write "<select id='repair_type' name='repair_type' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If LCase(repair_type) = "none" Then
			Response.Write "<option value='None' selected>None"
		Else
			Response.Write "<option value='None'>None"
		End If
		If LCase(repair_type) = "major" Then
			Response.Write "<option value='Major' selected>Major"
		Else
			Response.Write "<option value='Major'>Major"
		End If
		If LCase(repair_type) = "minor" Then
			Response.Write "<option value='Minor' selected>Minor"
		Else
			Response.Write "<option value='Minor'>Minor"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write "<td id='smalltd'>Type:</td>"
		Response.Write "<td id='smalltd'>"
		Response.Write repair_type & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	If editMode = True Then
		Response.Write "<td id='formtd'>Policy Verify:</td>"
		Response.Write "<td id='formtd'>"
		Response.Write "<select id='policy_verify' name='policy_verify' onchange='setupdate();'>"
		Response.Write "<option value=''>"
		If LCase(policy_verify) = "yes" Then
			Response.Write "<option value='Yes' selected>Yes"
		Else
			Response.Write "<option value='Yes'>Yes"
		End If
		If LCase(policy_verify) = "no" Then
			Response.Write "<option value='No' selected>No"
		Else
			Response.Write "<option value='No'>No"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write "<td id='smalltd'>Policy Verify:</td>"
		Response.Write "<td id='smalltd'>"
		Response.Write policy_verify & "</td>"
	End If
	Response.Write "<td id='smalltd' colspan='2'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "</td>"
	Response.Write "<td id='formtd' style='width:55%'>"
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	If editMode = True Then
		Response.Write "<td id='formtd' style='width:31%'>Inspection Company:</td>"
		Response.Write "<td id='formtd' colspan='4'>"
		Response.Write "<input type='text' class='text' style='width:100%' id='inspection_company' name='inspection_company' value='" & inspection_company & "' onchange='setupdate();' /></td>"
	Else
		Response.Write "<td id='smalltd' style='width:31%'>Inspection Company:</td>"
		Response.Write "<td id='smalltd' colspan='4'>"
		Response.Write inspection_company & "</td>"
	End If
	Response.Write "<tr>"
	If editMode = True Then
		Response.Write "<td id='formtd' style='width:31%'>Inspector:</td>"
		Response.Write "<td id='formtd' colspan='4'>"
		Response.Write "<input type='text' class='text' style='width:100%' id='inspected_by' name='inspected_by' value='" & inspected_by & "' onchange='setupdate();' /></td>"
	Else
		Response.Write "<td id='smalltd' style='width:31%'>Inspector:</td>"
		Response.Write "<td id='smalltd' colspan='4'>"
		Response.Write inspected_by & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	If editMode = True Then
		Response.Write "<td id='formtd'>Next Inspection Date:</td>"
		Response.Write "<td id='formtd' colspan='4'>"
		Response.Write "<input type='text' class='text' id='next_inspection_due' name='next_inspection_due' value='" & next_inspection_due & "' onchange='chkDate(this);setupdate();' />"
		Response.Write "<a href='javascript: void(0);' onclick='displayDatePicker(""next_inspection_due"");setupdate();return false;'><img src='../images/calendar.bmp' alt='Calendar' style='vertical-align:top'></a></td>"
	Else
		Response.Write "<td id='smalltd'>Next Inspection Date:</td>"
		Response.Write "<td id='smalltd' colspan='4'>"
		Response.Write next_inspection_due & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	If editMode = True Then
		Response.Write "<td id='formtd'>Previous Inspection:</td>"
		Response.Write "<td id='formtd' style='width:21%'>" & previous_inspection_date & "</td>"
		Response.Write "<td id='formtd' style='width:23%'>Set Frequency:</td>"
		Response.Write "<td id='formtd' style='width:11%'>"
		If UCase(previous_inspection_date) = "NONE" Or previous_inspection_date = "" Then
			previous_inspection_date = Date
		End If
		Response.Write "<input type='text' class='text' style='width:100%' id='set_frequency' name='set_frequency' value='" & set_frequency & "' onchange='chkNumeric(this);addDate(""" & previous_inspection_date & """,document.form1.set_frequency.value,document.form1.set_frequency_units.value);return false;' /></td>"
	Else
		Response.Write "<td id='smalltd'>Previous Inspection:</td>"
		Response.Write "<td id='smalltd' style='width:21%'>" & previous_inspection_date & "</td>"
		Response.Write "<td id='smalltd' style='width:23%'>Set Frequency:</td>"
		Response.Write "<td id='smalltd' style='width:11%'>"
		Response.Write set_frequency & "</td>"
	End If
	If editMode = True Then
		Response.Write "<td id='formtd' style='width:13%'>"
		Response.Write "<select style='width:100%' id='set_frequency_units' name='set_frequency_units' onchange='addDate(""" & previous_inspection_date & """,document.form1.set_frequency.value,document.form1.set_frequency_units.value);return false;'>"
		Response.Write "<option value=''>"
		If set_frequency_units = "days" Then
			Response.Write "<option value='days' selected>days"
		Else
			Response.Write "<option value='days'>days"
		End If
		If set_frequency_units = "months" Then
			Response.Write "<option value='months' selected>months"
		Else
			Response.Write "<option value='months'>months"
		End If
		If set_frequency_units = "years" Then
			Response.Write "<option value='years' selected>years"
		Else
			Response.Write "<option value='years'>years"
		End If
		Response.Write "</select></td>"
	Else
		Response.Write "<td id='smalltd' style='width:13%'>"
		Response.Write set_frequency_units & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='2'>COMMENTS:</td>"
	Response.Write "</tr>"
	Response.Write "<td id='formtd' colspan='2'>"
	If editMode = True Then
		Response.Write "<textarea style='width:100%;height:40px' id='comments' name='comments' rows='2' cols='80' onchange='setupdate();'>" & comments & "</textarea></td>"
	Else
		Response.Write comments & "</td>"
	End If
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
	Response.Write "<td id='formtd' colspan='2'>"
	If editMode = True Then
		Response.Write "<textarea style='width:100%;height:40px' id='discrepency_comments' name='discrepency_comments' rows='2' cols='80' onchange='setupdate();'>" & discrepency_comments & "</textarea></td>"
	Else
		Response.Write discrepency_comments & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2' style='font-weight:bold'>&nbsp;&nbsp;FOLLOW-UP:</td>"
	Response.Write "</tr>"
	Response.Write "<td id='formtd' colspan='2'>"
	If editMode = True Then
		Response.Write "<textarea style='width:100%;height:40px' id='discrepency_followup' name='discrepency_followup' rows='2' cols='80' onchange='setupdate();'>" & discrepency_followup & "</textarea></td>"
	Else
		Response.Write discrepency_followup & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	
	Response.Write "</td></tr>"
	Response.Write "<tr><td>"
	Response.Write "<div style='text-align:left'>"
	Response.Write "<table style='width:50%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2' style='font-weight:bold'>Reference:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:45%;font-weight:bold'>Location</td>"
	Response.Write "<td id='formtd' style='width:40%;font-weight:bold'>Reading</td>"
'	If editMode = True Then
'		Response.Write "<td id='formtd' class='noprint' style='width:15%'>"
'		Response.Write "<a href='javascript:addReading();' title='Add a reading record'>Add New</a></td>"
'	End If
	Response.Write "<td id='formtd' style='width:15%'>&nbsp;</td>"
	Response.Write "</tr>"
	'If this is an existing inspection, retrieve and display the existing
	'readings.
	If Request("inspectionID") <> "" Then
		sqlString = "SELECT * FROM tank_inspection_readings " & _
				"WHERE inspection_data_id=" & Request("inspectionID")
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			Do While Not rs.EOF
				Response.Write "<tr>"
				If Request("readingID") <> "" And editMode = True Then
					If CLng(Request("readingID")) = CLng(rs("inspection_reading_id")) Then
						Response.Write "<td id='formtd'>"
						Response.Write "<input type='text' class='text' style='width:100%' id='edit_location' name='edit_location' value='" & rs("location") & "' /></td>"
						Response.Write "<td id='formtd'>"
						Response.Write "<input type='text' class='text' style='width:100%' id='edit_reading' name='edit_reading' value='" & rs("reading") & "' /></td>"
						Response.Write "<td id='formtd'>"
						Response.Write "<a href='javascript:updateReading(" & rs("inspection_reading_id") & ");' title='Update this record in the database'>Submit</a></td>"
					Else
						Response.Write "<td id='formtd'>" & rs("location") & "</td>"
						Response.Write "<td id='formtd'>" & rs("reading") & "</td>"
						Response.Write "<td id='formtd'>"
						Response.Write "<a href='javascript:editReading(" & rs("inspection_reading_id") & ");' title='Open this record for editing'>Edit</a></td>"
					End If
				Else
					Response.Write "<td id='formtd'>" & rs("location") & "</td>"
					Response.Write "<td id='formtd'>" & rs("reading") & "</td>"
					If editMode = True Then
						Response.Write "<td id='formtd'>"
						Response.Write "<a href='javascript:editReading(" & rs("inspection_reading_id") & ");' title='Open this record for editing'>Edit</a></td>"
					End If
				End If
				Response.Write "</tr>"
				rs.MoveNext
			Loop
		End If
		rs.Close
	End If
	'If the reading id = -1, draw fields to allow the user to enter a new reading.
'	If Request("readingID") <> "" Then
'		If CLng(Request("readingID")) = -1 Then
	'If the reading id doesn't exist, draw fields to allow the user to enter a new reading.
	If editMode And (Request("readingID") = "" Or Request("readingID") = "-1") Then
			Response.Write "<tr>"
			Response.Write "<td id='formtd'>"
			Response.Write "<input type='text' class='text' style='width:100%' id='edit_location' name='edit_location' value='' /></td>"
			Response.Write "<td id='formtd'>"
			Response.Write "<input type='text' class='text' style='width:100%' id='edit_reading' name='edit_reading' value='' /></td>"
			Response.Write "<td id='formtd'>"
			Response.Write "<a href='javascript:updateReading(-1);' title='Insert this record into the database'>Submit</a></td>"
			Response.Write "</tr>"
	End If
'		End If
'	End If
	
	Response.Write "</table>"
	Response.Write "</div>"
		
	If editMode = True Then
		Response.Write "<div class='noprint' style='text-align:center'>"
		Response.Write "<br />"
		Response.Write "<hr />"
		Response.Write "<button type='button' id='submit1' name='submit1' onclick='saveData();'>Submit</button>"
		Response.Write "</div>"
	End If

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
<script language="VBScript">
<!--
Function checkDate_onchange(index)
	Dim strDate
	On Error Resume Next
	If index = 0 Then
 		strDate = document.form1.start_date.value
 		strDate = FormatDateTime(strDate,vbShortDate)
	ElseIf index = 1 Then
 		strDate = document.form1.end_date.value
 		strDate = FormatDateTime(strDate,vbShortDate)
 	End If
	If Err <> 0 Then
		MsgBox "Invalid date format entered: " & strDate
	End If
End Function
//-->
</script>
</html>
