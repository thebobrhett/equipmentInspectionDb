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
 window.open("http://mogsb8/inspections/cv_technicaldata.asp?itemID="+document.form1.itemID.value+"&edit=false","TechnicalData");
}
function saveData() {
 needToConfirm=false;
 document.form1.submit();
}
function setupdate() {
 needToConfirm=true;
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
<title>PSV Inspection</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="InspectionsFunctions.asp"-->
<link rel=STYLESHEET href='formstyle.css' type='text/css'>
</head>
<%
'*************
' Revision History
' 
' Keith Brooks - Tuesday, March 15, 2011
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
'Inspection items
Dim inspection_date
Dim vent_nameplate_matches
Dim vent_decontaminated
Dim flame_arr_nameplate_matches
Dim flame_arr_decontaminated
Dim vent_inlet_condition
Dim vent_inlet_requires_cleaning
Dim vent_piping_condition
Dim vent_piping_requires_cleaning
Dim vent_other
Dim vent_body
Dim flame_arrester_condition
Dim flame_arrester_requires_cleaning
Dim flame_arrester_requires_repair
Dim padding_regulator_condition
Dim padding_regulator_gauge_condition
Dim replace_regulator
Dim replace_gauge
Dim pressure_pallet_condition
Dim pressure_pallet_requires_cleaning
Dim pressure_pallet_requires_repair
Dim pressure_pallet_operated_manually
Dim vacuum_pallet_condition
Dim vacuum_pallet_requires_cleaning
Dim vacuum_pallet_requires_repair
Dim vacuum_pallet_operated_manually
Dim inspection_company
Dim inspected_by
Dim inspected_date
Dim pressure_pallet_cleaned
Dim pressure_pallet_seats_replaced
Dim pressure_pallet_guides_replaced
Dim pressure_pallet_other_repairs
Dim vacuum_pallet_cleaned
Dim vacuum_pallet_seats_replaced
Dim vacuum_pallet_guides_replaced
Dim vacuum_pallet_other_repairs
Dim flame_arrester_screen_cleaned
Dim conservation_vent_replaced
Dim flame_arrester_replaced
Dim serial_number_new_vent
Dim serial_number_new_flame_arrester
Dim work_order_number
Dim next_inspection_due
Dim previous_inspection
Dim set_frequency
Dim set_frequency_units
Dim policy_insp_verify
Dim repair_required
Dim repair_type
Dim repair_performed
Dim regulator_repaired
Dim regulator_gauge_repaired
Dim regulator_repaired_set_point
Dim regulator_repaired_set_point_units
Dim regulator_repaired_range_from
Dim regulator_repaired_range_to
Dim regulator_repaired_range_units
Dim repair_company
Dim repaired_by
Dim repaired_date
Dim cleaned_by
Dim cleaned_date
Dim flange_bolts_replaced
Dim flange_bolts_replaced_type
Dim flange_bolts_torqued
Dim flange_bolts_torqued_type
Dim regulator_replaced
Dim regulator_gauge_replaced
Dim regulator_replaced_set_point
Dim regulator_replaced_set_point_units
Dim regulator_replaced_range_from
Dim regulator_replaced_range_to
Dim regulator_replaced_range_units
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
access = UserAccess("inspections", "psv_inspection", currentuser)
If access <> "none" Then

	If LCase(Request("edit")) = "true" And (access = "write" Or access = "delete") Then
		editMode = True
		field_disabled = ""
		Response.Write "<body  onload='document.form1.inspection_date.focus();'>"
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
	Response.Write "<input type='hidden' id='equipType' name='equipType' value='psv' />"
	
	'If this is an existing inspection, get the record from the database.
	If Request("inspectionID") <> "" Then
		'Save the value of the inspection id.
		Response.Write "<input type='hidden' id='inspectionID' name='inspectionID' value='" & Request("inspectionID") & "' />"

		'Get the existing inspection data.
		sqlString = "SELECT * FROM psv_inspection_data " & _
				"WHERE inspection_data_id=" & Request("inspectionID")
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			itemID = rs("equipment_item_id")

			'Save the value of the item id.
			Response.Write "<input type='hidden' id='itemID' name='itemID' value='" & itemID & "' />"

			'Get the static technical data for this item from the database.
			sqlString = "SELECT equipment_item_tag,equipment_item_name,assembly,area,b.* " & _
					"FROM equipment_items a INNER JOIN psv_technical_data b " & _
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
				manufacturer = rs2("manufacturer")
				model_number = rs2("model_number")
				vacuum_set_point = rs2("vacuum_set_point")
				vacuum_set_point_units = rs2("vacuum_set_point_units")
				pressure_set_point = rs2("pressure_set_point")
				pressure_set_point_units = rs2("pressure_set_point_units")
				pad_set_point = rs2("pad_set_point")
				pad_set_point_units = rs2("pad_set_point_units")
				reg_gauge_range_from = rs2("reg_gauge_range_from")
				reg_gauge_range_to = rs2("reg_gauge_range_to")
				reg_gauge_range_units = rs2("reg_gauge_range_units")
				fluid_service = rs2("fluid_service")
				specification_number = rs2("specification_number")
				serial_number = rs2("serial_number")
				arrester_manufacturer = rs2("arrester_manufacturer")
				arrester_model_number = rs2("arrester_model_number")
				arrester_serial_number = rs2("arrester_serial_number")
				arrester_spec_number = rs2("arrester_spec_number")
				fluid_state = rs2("fluid_state")
			End If
			rs2.Close
			Set rs2 = Nothing
			
			'Fill in the existing inspection data variables.
			If IsNull(rs("inspection_date")) Then
				inspection_date = ""
			Else
				inspection_date = FormatDateTime(rs("inspection_date"),2)
			End If
			vent_nameplate_matches = rs("vent_nameplate_matches")
			vent_decontaminated = rs("vent_decontaminated")
			flame_arr_nameplate_matches = rs("flame_arr_nameplate_matches")
			flame_arr_decontaminated = rs("flame_arr_decontaminated")
			vent_inlet_condition = rs("vent_inlet_condition")
			vent_inlet_requires_cleaning = rs("vent_inlet_requires_cleaning")
			vent_piping_condition = rs("vent_piping_condition")
			vent_piping_requires_cleaning = rs("vent_piping_requires_cleaning")
			vent_other = rs("vent_other")
			vent_body = rs("vent_body")
			flame_arrester_condition = rs("flame_arrester_condition")
			flame_arrester_requires_cleaning = rs("flame_arrester_requires_cleaning")
			flame_arrester_requires_repair = rs("flame_arrester_requires_repair")
			padding_regulator_condition = rs("padding_regulator_condition")
			padding_regulator_gauge_condition = rs("padding_regulator_gauge_condition")
			replace_regulator = rs("replace_regulator")
			replace_gauge = rs("replace_gauge")
			pressure_pallet_condition = rs("pressure_pallet_condition")
			pressure_pallet_requires_cleaning = rs("pressure_pallet_requires_cleaning")
			pressure_pallet_requires_repair = rs("pressure_pallet_requires_repair")
			pressure_pallet_operated_manually = rs("pressure_pallet_operated_manually")
			vacuum_pallet_condition = rs("vacuum_pallet_condition")
			vacuum_pallet_requires_cleaning = rs("vacuum_pallet_requires_cleaning")
			vacuum_pallet_requires_repair = rs("vacuum_pallet_requires_repair")
			vacuum_pallet_operated_manually = rs("vacuum_pallet_operated_manually")
			inspection_company = rs("inspection_company")
			inspected_by = rs("inspected_by")
			If IsNull(rs("inspected_date")) Then
				inspected_date = ""
			Else
				inspected_date = FormatDateTime(rs("inspected_date"),2)
			End If
			pressure_pallet_cleaned = rs("pressure_pallet_cleaned")
			pressure_pallet_seats_replaced = rs("pressure_pallet_seats_replaced")
			pressure_pallet_guides_replaced = rs("pressure_pallet_guides_replaced")
			pressure_pallet_other_repairs = rs("pressure_pallet_other_repairs")
			vacuum_pallet_cleaned = rs("vacuum_pallet_cleaned")
			vacuum_pallet_seats_replaced = rs("vacuum_pallet_seats_replaced")
			vacuum_pallet_guides_replaced = rs("vacuum_pallet_guides_replaced")
			vacuum_pallet_other_repairs = rs("vacuum_pallet_other_repairs")
			flame_arrester_screen_cleaned = rs("flame_arrester_screen_cleaned")
			conservation_vent_replaced = rs("conservation_vent_replaced")
			flame_arrester_replaced = rs("flame_arrester_replaced")
			serial_number_new_vent = rs("serial_number_new_vent")
			serial_number_new_flame_arrester = rs("serial_number_new_flame_arrester")
			work_order_number = rs("work_order_number")
			next_inspection_due = rs("next_inspection_due")
			previous_inspection = rs("previous_inspection")
			set_frequency = rs("set_frequency")
			set_frequency_units = rs("set_frequency_units")
			policy_insp_verify = rs("policy_insp_verify")
			repair_required = rs("repair_required")
			repair_type = rs("repair_type")
			repair_performed = rs("repair_performed")
			regulator_repaired = rs("regulator_repaired")
			regulator_gauge_repaired = rs("regulator_gauge_repaired")
			regulator_repaired_set_point = rs("regulator_repaired_set_point")
			regulator_repaired_set_point_units = rs("regulator_repaired_set_point_units")
			regulator_repaired_range_from = rs("regulator_repaired_range_from")
			regulator_repaired_range_to = rs("regulator_repaired_range_to")
			regulator_repaired_range_units = rs("regulator_repaired_range_units")
			repair_company = rs("repair_company")
			repaired_by = rs("repaired_by")
			If IsNull(rs("repaired_date")) Then
				repaired_date = ""
			Else
				repaired_date = FormatDateTime(rs("repaired_date"),2)
			End If
			cleaned_by = rs("cleaned_by")
			cleaned_date = rs("cleaned_date")
			flange_bolts_replaced = rs("flange_bolts_replaced")
			flange_bolts_replaced_type = rs("flange_bolts_replaced_type")
			flange_bolts_torqued = rs("flange_bolts_torqued")
			flange_bolts_torqued_type = rs("flange_bolts_torqued_type")
			regulator_replaced = rs("regulator_replaced")
			regulator_gauge_replaced = rs("regulator_gauge_replaced")
			regulator_replaced_set_point = rs("regulator_replaced_set_point")
			regulator_replaced_set_point_units = rs("regulator_replaced_set_point_units")
			regulator_replaced_range_from = rs("regulator_replaced_range_from")
			regulator_replaced_range_to = rs("regulator_replaced_range_to")
			regulator_replaced_range_units = rs("regulator_replaced_range_units")
			installed_by = rs("installed_by")
			installed_date = rs("installed_date")
			comment = rs("comment")
			discrepency_comments = rs("discrepency_comments")
			discrepency_followup = rs("discrepency_followup")
			
		End If
	Else
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
			
			'Initialize the inspection data variables.
			inspection_date = Date
			vent_nameplate_matches = 0
			vent_decontaminated = 0
			flame_arr_nameplate_matches = 0
			flame_arr_decontaminated = 0
			vent_inlet_condition = ""
			vent_inlet_requires_cleaning = 0
			vent_piping_condition = ""
			vent_piping_requires_cleaning = 0
			vent_other = ""
			vent_body = ""
			flame_arrester_condition = ""
			flame_arrester_requires_cleaning = 0
			flame_arrester_requires_repair = 0
			padding_regulator_condition = ""
			padding_regulator_gauge_condition = ""
			replace_regulator = 0
			replace_gauge = 0
			pressure_pallet_condition = ""
			pressure_pallet_requires_cleaning = 0
			pressure_pallet_requires_repair = 0
			pressure_pallet_operated_manually = 0
			vacuum_pallet_condition = ""
			vacuum_pallet_requires_cleaning = 0
			vacuum_pallet_requires_repair = 0
			vacuum_pallet_operated_manually = 0
			inspection_company = ""
			inspected_by = ""
			inspected_date = Date
			pressure_pallet_cleaned = 0
			pressure_pallet_seats_replaced = 0
			pressure_pallet_guides_replaced = 0
			pressure_pallet_other_repairs = ""
			vacuum_pallet_cleaned = 0
			vacuum_pallet_seats_replaced = 0
			vacuum_pallet_guides_replaced = 0
			vacuum_pallet_other_repairs = ""
			flame_arrester_screen_cleaned = 0
			conservation_vent_replaced = 0
			flame_arrester_replaced = 0
			serial_number_new_vent = ""
			serial_number_new_flame_arrester = ""
			work_order_number = ""
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
			policy_insp_verify = "no"
			repair_required = "no"
			repair_type = "none"
			repair_performed = "no"
			regulator_repaired = 0
			regulator_gauge_repaired = 0
			regulator_repaired_set_point = ""
			regulator_repaired_set_point_units = ""
			regulator_repaired_range_from = ""
			regulator_repaired_range_to = ""
			regulator_repaired_range_units = ""
			repair_company = ""
			repaired_by = ""
			repaired_date = ""
			cleaned_by = ""
			cleaned_date = ""
			flange_bolts_replaced = 0
			flange_bolts_replaced_type = ""
			flange_bolts_torqued = 0
			flange_bolts_torqued_type = ""
			regulator_replaced = 0
			regulator_gauge_replaced = 0
			regulator_replaced_set_point = ""
			regulator_replaced_set_point_units = ""
			regulator_replaced_range_from = ""
			regulator_replaced_range_to = ""
			regulator_replaced_range_units = ""
			installed_by = ""
			installed_date = ""
			comment = ""
			discrepency_comments = ""
			discrepency_followup = ""
		End If
	End If
	
	'Save the previous inspection date in a hidden tag so it can be saved with
	'the rest of the data.
	Response.Write "<input type='hidden' id='previous_inspection' name='previous_inspection' value='" & previous_inspection & "' />"

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
	
	Response.Write "<tbody valign='top'><tr><td>"
'	Response.Write "<br />"
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='5'>CONSERVATION VENT DESIGN:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='techtd' style='width:23%'>Manufacturer: </td>"
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
	Response.Write "<td id='techtd' style='width:17%'>" & vacuum_set_point_units & "</td>"
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
	Response.Write "</table>"

	'Draw the button to open the technical data form.
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='4' style='text-align:right'>"
	Response.Write "<input type='button' class='noprint' id='techdata' name='techdata' value='Technical Data' onclick='opentechdata();return false;' /></td>"
	Response.Write "</tr>"
	
	'Draw the data entry section.
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='4'>INSPECTION:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:35%'>Vent Nameplate Data Match Above Data: </td>"
	Response.Write "<td id='formtd' style='width:15%'>"
	If vent_nameplate_matches = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='vent_nameplate_matches' name='vent_nameplate_matches' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='vent_nameplate_matches' name='vent_nameplate_matches' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd' style='width:25%'>Vent Decontaminated: </td>"
	Response.Write "<td id='formtd' style='width:25%'>"
	If vent_decontaminated = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='vent_decontaminated' name='vent_decontaminated' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='vent_decontaminated' name='vent_decontaminated' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Flame Arr Nameplate Match Above Data: </td>"
	Response.Write "<td id='formtd'>"
	If flame_arr_nameplate_matches = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='flame_arr_nameplate_matches' name='flame_arr_nameplate_matches' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='flame_arr_nameplate_matches' name='flame_arr_nameplate_matches' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Flame Arr Decontaminated: </td>"
	Response.Write "<td id='formtd'>"
	If flame_arr_decontaminated = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='flame_arr_decontaminated' name='flame_arr_decontaminated' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='flame_arr_decontaminated' name='flame_arr_decontaminated' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr><td id='formtd' colspan='4'>&nbsp;</td></tr>"
	Response.Write "</table>"
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='2' style='text-align:center'>VENT INLET</td>"
	Response.Write "<td id='grouptd' colspan='4' style='text-align:center'>VENT PIPING</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:23%'>Condition: </td>"
	Response.Write "<td id='formtd' style='width:27%'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='vent_inlet_condition' name='vent_inlet_condition' value='" & vent_inlet_condition & "' onchange='setupdate();' /></td>"
	Else
		Response.Write vent_inlet_condition & "</td>"
	End If
	Response.Write "<td id='formtd' style='width:20%'>Condition: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='vent_piping_condition' name='vent_piping_condition' value='" & vent_piping_condition & "' onchange='setupdate();' /></td>"
	Else
		Response.Write vent_piping_condition & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Requires Cleaning: </td>"
	Response.Write "<td id='formtd'>"
	If vent_inlet_requires_cleaning = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='vent_inlet_requires_cleaning' name='vent_inlet_requires_cleaning' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='vent_inlet_requires_cleaning' name='vent_inlet_requires_cleaning' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Requires Cleaning: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	If vent_piping_requires_cleaning = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='vent_piping_requires_cleaning' name='vent_piping_requires_cleaning' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='vent_piping_requires_cleaning' name='vent_piping_requires_cleaning' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='vertical-align:top'>Other: </td>"
	Response.Write "<td id='formtd' colspan='5'>"
	If editMode = True Then
		Response.Write "<textarea style='width:100%;height:40px' id='vent_other' name='vent_other' rows='2' cols='80' onchange='setupdate();'>" & vent_other & "</textarea></td>"
	Else
		Response.Write vent_other & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Body: </td>"
	Response.Write "<td id='formtd' colspan='5'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='vent_body' name='vent_body' value='" & vent_body & "' onchange='setupdate();' /></td>"
	Else
		Response.Write vent_body & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr><td id='formtd' colspan='6'>&nbsp;</td></tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='2' style='text-align:center'>FLAME ARRESTER/SCREEN</td>"
	Response.Write "<td id='grouptd' colspan='4' style='text-align:center'>PADDING REGULATOR</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Condition: </td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='flame_arrester_condition' name='flame_arrester_condition' value='" & flame_arrester_condition & "' onchange='setupdate();' /></td>"
	Else
		Response.Write flame_arrester_condition & "</td>"
	End If
	Response.Write "<td id='formtd'>Condition: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='padding_regulator_condition' name='padding_regulator_condition' value='" & padding_regulator_condition & "' onchange='setupdate();' /></td>"
	Else
		Response.Write padding_regulator_condition & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Requires Cleaning:</td>"
	Response.Write "<td id='formtd'>"
	If flame_arrester_requires_cleaning = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='flame_arrester_requires_cleaning' name='flame_arrester_requires_cleaning' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='flame_arrester_requires_cleaning' name='flame_arrester_requires_cleaning' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Gauge Condition:</td>"
	Response.Write "<td id='formtd' colspan='3'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='padding_regulator_gauge_condition' name='padding_regulator_gauge_condition' value='" & padding_regulator_gauge_condition & "' onchange='setupdate();' /></td>"
	Else
		Response.Write padding_regulator_gauge_condition & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Requires Repair:</td>"
	Response.Write "<td id='formtd'>"
	If flame_arrester_requires_repair = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='flame_arrester_requires_repair' name='flame_arrester_requires_repair' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='flame_arrester_requires_repair' name='flame_arrester_requires_repair' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Replace Regulator:</td>"
	Response.Write "<td id='formtd' style='width:10%'>"
	If replace_regulator = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='replace_regulator' name='replace_regulator' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='replace_regulator' name='replace_regulator' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd' style='width:15%'>Replace Gauge:</td>"
	Response.Write "<td id='formtd' style='width:5%'>"
	If replace_gauge = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='replace_gauge' name='replace_gauge' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='replace_gauge' name='replace_gauge' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr><td id='formtd' colspan='6'>&nbsp;</td></tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='2' style='text-align:center'>PRESSURE PALLET</td>"
	Response.Write "<td id='grouptd' colspan='4' style='text-align:center'>VACUUM PALLET</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Condition: </td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='pressure_pallet_condition' name='pressure_pallet_condition' value='" & pressure_pallet_condition & "' onchange='setupdate();' /></td>"
	Else
		Response.Write pressure_pallet_condition & "</td>"
	End If
	Response.Write "<td id='formtd'>Condition: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='vacuum_pallet_condition' name='vacuum_pallet_condition' value='" & vacuum_pallet_condition & "' onchange='setupdate();' /></td>"
	Else
		Response.Write vacuum_pallet_condition & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Requires Cleaning: </td>"
	Response.Write "<td id='formtd'>"
	If pressure_pallet_requires_cleaning = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='pressure_pallet_requires_cleaning' name='pressure_pallet_requires_cleaning' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='pressure_pallet_requires_cleaning' name='pressure_pallet_requires_cleaning' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Requires Cleaning: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	If vacuum_pallet_requires_cleaning = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='vacuum_pallet_requires_cleaning' name='vacuum_pallet_requires_cleaning' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='vacuum_pallet_requires_cleaning' name='vacuum_pallet_requires_cleaning' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Requires Repair: </td>"
	Response.Write "<td id='formtd'>"
	If pressure_pallet_requires_repair = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='pressure_pallet_requires_repair' name='pressure_pallet_requires_repair' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='pressure_pallet_requires_repair' name='pressure_pallet_requires_repair' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Requires Repair: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	If vacuum_pallet_requires_repair = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='vacuum_pallet_requires_repair' name='vacuum_pallet_requires_repair' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='vacuum_pallet_requires_repair' name='vacuum_pallet_requires_repair' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Operated Manually: </td>"
	Response.Write "<td id='formtd'>"
	If pressure_pallet_operated_manually = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='pressure_pallet_operated_manually' name='pressure_pallet_operated_manually' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='pressure_pallet_operated_manually' name='pressure_pallet_operated_manually' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Operated Manually: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	If vacuum_pallet_operated_manually = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='vacuum_pallet_operated_manually' name='vacuum_pallet_operated_manually' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='vacuum_pallet_operated_manually' name='vacuum_pallet_operated_manually' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr><td id='formtd' colspan='6'>&nbsp;</td></tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='font-weight:bold'>Inspection Company: </td>"
	Response.Write "<td id='formtd' colspan='5'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='inspection_company' name='inspection_company' value='" & inspection_company & "' onchange='setupdate();' /></td>"
	Else
		Response.Write inspection_company & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='font-weight:bold'>Field Inspection By: </td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='inspected_by' name='inspected_by' value='" & inspected_by & "' onchange='setupdate();' /></td>"
	Else
		Response.Write inspected_by & "</td>"
	End If
	Response.Write "<td id='formtd' style='font-weight:bold'>Date: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' id='inspected_date' name='inspected_date' value='" & inspected_date & "' onchange='chkDate(this);setupdate();' />"
		Response.Write "<a href='javascript: void(0);' onclick='displayDatePicker(""inspected_date"");setupdate();return false;'><img src='../images/calendar.bmp' alt='Calendar' style='vertical-align:top'></a></td>"
	Else
		Response.Write inspected_date & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr><td id='formtd' colspan='6'>&nbsp;</td></tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='4'>REPAIRS MADE:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='2' style='text-align:center'>PRESSURE PALLET</td>"
	Response.Write "<td id='grouptd' colspan='4' style='text-align:center'>VACUUM PALLET</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Cleaned: </td>"
	Response.Write "<td id='formtd'>"
	If pressure_pallet_cleaned = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='pressure_pallet_cleaned' name='pressure_pallet_cleaned' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='pressure_pallet_cleaned' name='pressure_pallet_cleaned' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Cleaned: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	If vacuum_pallet_cleaned = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='vacuum_pallet_cleaned' name='vacuum_pallet_cleaned' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='vacuum_pallet_cleaned' name='vacuum_pallet_cleaned' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Seats Replaced: </td>"
	Response.Write "<td id='formtd'>"
	If pressure_pallet_seats_replaced = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='pressure_pallet_seats_replaced' name='pressure_pallet_seats_replaced' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='pressure_pallet_seats_replaced' name='pressure_pallet_seats_replaced' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Seats Replaced: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	If vacuum_pallet_seats_replaced = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='vacuum_pallet_seats_replaced' name='vacuum_pallet_seats_replaced' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='vacuum_pallet_seats_replaced' name='vacuum_pallet_seats_replaced' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Guides Replaced: </td>"
	Response.Write "<td id='formtd'>"
	If pressure_pallet_guides_replaced = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='pressure_pallet_guides_replaced' name='pressure_pallet_guides_replaced' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='pressure_pallet_guides_replaced' name='pressure_pallet_guides_replaced' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Guides Replaced: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	If vacuum_pallet_guides_replaced = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='vacuum_pallet_guides_replaced' name='vacuum_pallet_guides_replaced' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='vacuum_pallet_guides_replaced' name='vacuum_pallet_guides_replaced' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='vertical-align:top'>Other: </td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<textarea style='width:90%;height:40px' id='pressure_pallet_other_repairs' name='pressure_pallet_other_repairs' rows='2' cols='40' onchange='setupdate();'>" & pressure_pallet_other_repairs & "</textarea></td>"
	Else
		Response.Write pressure_pallet_other_repairs & "</td>"
	End If
	Response.Write "<td id='formtd' style='vertical-align:top'>Other: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	If editMode = True Then
		Response.Write "<textarea style='width:100%;height:40px' id='vacuum_pallet_other_repairs' name='vacuum_pallet_other_repairs' rows='2' cols='40' onchange='setupdate();'>" & vacuum_pallet_other_repairs & "</textarea></td>"
	Else
		Response.Write vacuum_pallet_other_repairs & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:30%'>Flame Arrester Screen Cleaned: </td>"
	Response.Write "<td id='formtd' style='width:20%'>"
	If flame_arrester_screen_cleaned = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='flame_arrester_screen_cleaned' name='flame_arrester_screen_cleaned' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='flame_arrester_screen_cleaned' name='flame_arrester_screen_cleaned' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd' colspan='2'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Conservation Vent Replaced:</td>"
	Response.Write "<td id='formtd'>"
	If conservation_vent_replaced = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='conservation_vent_replaced' name='conservation_vent_replaced' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='conservation_vent_replaced' name='conservation_vent_replaced' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd' style='width:20%'>Serial No. New Vent:</td>"
	Response.Write "<td id='formtd' style='width:30%'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='serial_number_new_vent' name='serial_number_new_vent' value='" & serial_number_new_vent & "' onchange='setupdate();' /></td>"
	Else
		Response.Write serial_number_new_vent & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Flame Arrester Replaced:</td>"
	Response.Write "<td id='formtd'>"
	If flame_arrester_replaced = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='flame_arrester_replaced' name='flame_arrester_replaced' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='flame_arrester_replaced' name='flame_arrester_replaced' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Serial No. FA:</td>"
	Response.Write "<td id='formtd' style='width:30%'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='serial_number_new_flame_arrester' name='serial_number_new_flame_arrester' value='" & serial_number_new_flame_arrester & "' onchange='setupdate();' /></td>"
	Else
		Response.Write serial_number_new_flame_arrester & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='4'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2' style='padding:0px'>"
	Response.Write "<table style='width:100%;border:1px solid black'>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:40%'>Policy Insp. Verify:</td>"
	If editMode = True Then
		Response.Write "<td id='formtd' style='width:5%'>"
		If LCase(policy_insp_verify) = "yes" Then
			Response.Write "<input type='radio' class='radio' id='policy_insp_verify' name='policy_insp_verify' value='yes' " & field_disabled & " onchange='setupdate();' checked />yes</td>"
		Else
			Response.Write "<input type='radio' class='radio' id='policy_insp_verify' name='policy_insp_verify' value='yes' " & field_disabled & " onchange='setupdate();' />yes</td>"
		End If
		Response.Write "<td id='formtd' style='width:5%'>"
		If LCase(policy_insp_verify) = "no" Then
			Response.Write "<input type='radio' class='radio' id='policy_insp_verify' name='policy_insp_verify' value='no' " & field_disabled & " onchange='setupdate();' checked />no</td>"
		Else
			Response.Write "<input type='radio' class='radio' id='policy_insp_verify' name='policy_insp_verify' value='no' " & field_disabled & " onchange='setupdate();' />no</td>"
		End If
	Else
		Response.Write "<td id='formtd' style='width:10%'>"
		Response.Write policy_insp_verify & "</td>"
	End If
	Response.Write "<td id='formtd' style='width:40%'>Repair Performed:</td>"
	If editMode = True Then
		Response.Write "<td id='formtd' style='width:5%'>"
		If LCase(repair_performed) = "yes" Then
			Response.Write "<input type='radio' class='radio' id='repair_performed' name='repair_performed' value='yes' " & field_disabled & " onchange='setupdate();' checked />yes</td>"
		Else
			Response.Write "<input type='radio' class='radio' id='repair_performed' name='repair_performed' value='yes' " & field_disabled & " onchange='setupdate();' />yes</td>"
		End If
		Response.Write "<td id='formtd' style='width:5%'>"
		If LCase(repair_performed) = "no" Then
			Response.Write "<input type='radio' class='radio' id='repair_performed' name='repair_performed' value='no' " & field_disabled & " onchange='setupdate();' checked />no</td>"
		Else
			Response.Write "<input type='radio' class='radio' id='repair_performed' name='repair_performed' value='no' " & field_disabled & " onchange='setupdate();' />no</td>"
		End If
	Else
		Response.Write "<td id='formtd' style='width:10%'>"
		Response.Write repair_performed & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Repair Required:</td>"
	If editMode = True Then
		Response.Write "<td id='formtd'>"
		If LCase(repair_required) = "yes" Then
			Response.Write "<input type='radio' class='radio' id='repair_required' name='repair_required' value='yes' " & field_disabled & " onchange='setupdate();' checked />yes</td>"
		Else
			Response.Write "<input type='radio' class='radio' id='repair_required' name='repair_required' value='yes' " & field_disabled & " onchange='setupdate();' />yes</td>"
		End If
		Response.Write "<td id='formtd'>"
		If LCase(repair_required) = "no" Then
			Response.Write "<input type='radio' class='radio' id='repair_required' name='repair_required' value='no' " & field_disabled & " onchange='setupdate();' checked />no</td>"
		Else
			Response.Write "<input type='radio' class='radio' id='repair_required' name='repair_required' value='no' " & field_disabled & " onchange='setupdate();' />no</td>"
		End If
		Response.Write "<td id='formtd' colspan='3'>&nbsp;</td>"
	Else
		Response.Write "<td id='formtd'>"
		Response.Write repair_required & "</td>"
		Response.Write "<td id='formtd' colspan='2'>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Repair Type:</td>"
	If editMode = True Then
		Response.Write "<td id='formtd'>"
		If LCase(repair_type) = "none" Then
			Response.Write "<input type='radio' class='radio' id='repair_type' name='repair_type' value='none' " & field_disabled & " onchange='setupdate();' checked />none</td>"
		Else
			Response.Write "<input type='radio' class='radio' id='repair_type' name='repair_type' value='none' " & field_disabled & " onchange='setupdate();' />none</td>"
		End If
		Response.Write "<td id='formtd'>"
		If LCase(repair_type) = "major" Then
			Response.Write "<input type='radio' class='radio' id='repair_type' name='repair_type' value='major' " & field_disabled & " onchange='setupdate();' checked />major</td>"
		Else
			Response.Write "<input type='radio' class='radio' id='repair_type' name='repair_type' value='major' " & field_disabled & " onchange='setupdate();' />major</td>"
		End If
		Response.Write "<td id='formtd'>"
		If LCase(repair_type) = "minor" Then
			Response.Write "<input type='radio' class='radio' id='repair_type' name='repair_type' value='minor' " & field_disabled & " onchange='setupdate();' checked />minor</td>"
		Else
			Response.Write "<input type='radio' class='radio' id='repair_type' name='repair_type' value='minor' " & field_disabled & " onchange='setupdate();' />minor</td>"
		End If
		Response.Write "<td id='formtd' colspan='2'>&nbsp;</td>"
	Else
		Response.Write "<td id='formtd'>"
		Response.Write repair_type & "</td>"
		Response.Write "<td id='formtd' colspan='2'>"
	End If
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "</td>"
	Response.Write "<td id='formtd' colspan='2' style='padding:0px 0px 0px 5px'>"
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:39%'>Work Order Number:</td>"
	Response.Write "<td id='formtd' colspan='4'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='work_order_number' name='work_order_number' value='" & work_order_number & "' onchange='setupdate();' /></td>"
	Else
		Response.Write work_order_number & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Next Inspection Due:</td>"
	Response.Write "<td id='formtd' colspan='4'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' id='next_inspection_due' name='next_inspection_due' value='" & next_inspection_due & "' onchange='chkDate(this);setupdate();' />"
		Response.Write "<a href='javascript: void(0);' onclick='displayDatePicker(""next_inspection_due"");setupdate();return false;'><img src='../images/calendar.bmp' alt='Calendar' style='vertical-align:top'></a></td>"
	Else
		Response.Write next_inspection_due & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Previous Inspection:</td>"
	Response.Write "<td id='formtd' style='width:21%'>" & previous_inspection & "</td>"
	Response.Write "<td id='formtd' style='width:20%'>Set Freq:</td>"
	Response.Write "<td id='formtd' style='width:9%'>"
	If editMode = True Then
		If UCase(previous_inspection) = "NONE" Or previous_inspection = "" Then
			previous_inspection = Date
		End If
		Response.Write "<input type='text' class='text' style='width:100%' id='set_frequency' name='set_frequency' value='" & set_frequency & "' onchange='chkNumeric(this);addDate(""" & previous_inspection & """,document.form1.set_frequency.value,document.form1.set_frequency_units.value);return false;' /></td>"
	Else
		Response.Write set_frequency & "</td>"
	End If
	Response.Write "<td id='formtd' style='width:10%'>"
	If editMode = True Then
		Response.Write "<select style='width:70px' id='set_frequency_units' name='set_frequency_units' onchange='addDate(""" & previous_inspection & """,document.form1.set_frequency.value,document.form1.set_frequency_units.value);return false;'>"
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
		Response.Write set_frequency_units & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	
	Response.Write "<div style='page-break-before:always'></div>"
	
	Response.Write "<br />"
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='7'>REPAIRS MADE (cont):</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:23%'>Regulator Repaired:</td>"
	Response.Write "<td id='formtd' style='width:27%'>"
	If regulator_repaired = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='regulator_repaired' name='regulator_repaired' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='regulator_repaired' name='regulator_repaired' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd' style='width:20%'>Set Point:</td>"
	Response.Write "<td id='formtd' style='width:8%'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='regulator_repaired_set_point' name='regulator_repaired_set_point' value='" & regulator_repaired_set_point & "' onchange='chkNumeric(this);setupdate();' /></td>"
	Else
		Response.Write regulator_repaired_set_point & "</td>"
	End If
	Response.Write "<td id='formtd' colspan='3'>"
	If editMode = True Then
		Response.Write "(units) <input type='text' class='text' style='width:80%' id='regulator_repaired_set_point_units' name='regulator_repaired_set_points_units' value='" & regulator_repaired_set_point_units & "' onchange='setupdate();' /></td>"
	Else
		Response.Write regulator_repaired_set_point_units & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Regulator Gauge Repaired:</td>"
	Response.Write "<td id='formtd'>"
	If regulator_gauge_repaired = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='regulator_gauge_repaired' name='regulator_gauge_repaired' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='regulator_gauge_repaired' name='regulator_gauge_repaired' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Range:</td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='regulator_repaired_range_from' name='regulator_repaired_range_from' value='" & regulator_repaired_range_from & "' onchange='chkNumeric(this);setupdate();' /></td>"
	Else
		Response.Write regulator_repaired_range_from & "</td>"
	End If
	Response.Write "<td id='formtd' style='width:1%;text-align:center'>to</td>"
	Response.Write "<td id='formtd' style='width:8%'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='regulator_repaired_range_to' name='regulator_repaired_range_to' value='" & regulator_repaired_range_to & "' onchange='chkNumeric(this);setupdate();' /></td>"
	Else
		Response.Write regulator_repaired_range_to & "</td>"
	End If
	Response.Write "<td id='formtd' style='width:13%'>"
	If editMode = True Then
		Response.Write "(units) <input type='text' class='text' style='width:60%' id='regulator_repaired_range_units' name='regulator_repaired_range_units' value='" & regulator_repaired_range_units & "' onchange='setupdate();' /></td>"
	Else
		Response.Write regulator_repaired_range_units & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Repair Company: </td>"
	Response.Write "<td id='formtd' colspan='6'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='repair_company' name='repair_company' value='" & repair_company & "' onchange='setupdate();' /></td>"
	Else
		Response.Write repair_company & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Repaired By: </td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='repaired_by' name='repaired_by' value='" & repaired_by & "' onchange='setupdate();' /></td>"
	Else
		Response.Write repaired_by & "</td>"
	End If
	Response.Write "<td id='formtd'>Date: </td>"
	Response.Write "<td id='formtd' colspan='4'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' id='repaired_date' name='repaired_date' value='" & repaired_date & "' onchange='chkDate(this);setupdate();' />"
		Response.Write "<a href='javascript: void(0);' onclick='displayDatePicker(""repaired_date"");setupdate();return false;'><img src='../images/calendar.bmp' alt='Calendar' style='vertical-align:top'></a></td>"
	Else
		Response.Write repaired_date & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Cleaned By: </td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='cleaned_by' name='cleaned_by' value='" & cleaned_by & "' onchange='setupdate();' /></td>"
	Else
		Response.Write cleaned_by & "</td>"
	End If
	Response.Write "<td id='formtd'>Date: </td>"
	Response.Write "<td id='formtd' colspan='4'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' id='cleaned_date' name='cleaned_date' value='" & cleaned_date & "' onchange='chkDate(this);setupdate();' />"
		Response.Write "<a href='javascript: void(0);' onclick='displayDatePicker(""cleaned_date"");setupdate();return false;'><img src='../images/calendar.bmp' alt='Calendar' style='vertical-align:top'></a></td>"
	Else
		Response.Write cleaned_date & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr><td id='formtd' colspan='7'>&nbsp;</td></tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='7'>INSTALLATION:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Flange Bolts Replaced:</td>"
	Response.Write "<td id='formtd'>"
	If flange_bolts_replaced = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='flange_bolts_replaced' name='flange_bolts_replaced' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='flange_bolts_replaced' name='flange_bolts_replaced' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Type:</td>"
	Response.Write "<td id='formtd' colspan='4'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='flange_bolts_replaced_type' name='flange_bolts_replaced_type' value='" & flange_bolts_replaced_type & "' onchange='setupdate();' /></td>"
	Else
		Response.Write flange_bolts_replaced_type & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Flange Bolts Torqued:</td>"
	Response.Write "<td id='formtd'>"
	If flange_bolts_torqued = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='flange_bolts_torqued' name='flange_bolts_torqued' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='flange_bolts_torqued' name='flange_bolts_torqued' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Type:</td>"
	Response.Write "<td id='formtd' colspan='4'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='flange_bolts_torqued_type' name='flange_bolts_torqued_type' value='" & flange_bolts_torqued_type & "' onchange='setupdate();' /></td>"
	Else
		Response.Write flange_bolts_torqued_type & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Regulator Replaced:</td>"
	Response.Write "<td id='formtd'>"
	If regulator_replaced = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='regulator_replaced' name='regulator_replaced' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='regulator_replaced' name='regulator_replaced' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd' style='width:20%'>Set Point:</td>"
	Response.Write "<td id='formtd' style='width:8%'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='regulator_replaced_set_point' name='regulator_replaced_set_point' value='" & regulator_replaced_set_point & "' onchange='chkNumeric(this);setupdate();' /></td>"
	Else
		Response.Write regulator_replaced_set_point & "</td>"
	End If
	Response.Write "<td id='formtd' colspan='3'>"
	If editMode = True Then
		Response.Write "(units) <input type='text' class='text' style='width:80%' id='regulator_replaced_set_point_units' name='regulator_replaced_set_point_units' value='" & regulator_replaced_set_point_units & "' onchange='setupdate();' /></td>"
	Else
		Response.Write regulator_replaced_set_point_units & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Regulator Gauge Replaced:</td>"
	Response.Write "<td id='formtd'>"
	If regulator_gauge_replaced = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='regulator_gauge_replaced' name='regulator_gauge_replaced' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='regulator_gauge_replaced' name='regulator_gauge_replaced' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Range:</td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='regulator_replaced_range_from' name='regulator_replaced_range_from' value='" & regulator_replaced_range_from & "' onchange='chkNumeric(this);setupdate();' /></td>"
	Else
		Response.Write regulator_replaced_range_from & "</td>"
	End If
	Response.Write "<td id='formtd' style='width:1%;text-align:center'>to</td>"
	Response.Write "<td id='formtd' style='width:8%'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='regulator_replaced_range_to' name='regulator_replaced_range_to' value='" & regulator_replaced_range_to & "' onchange='chkNumeric(this);setupdate();' /></td>"
	Else
		Response.Write regulator_replaced_range_to & "</td>"
	End If
	Response.Write "<td id='formtd' style='width:13%'>"
	If editMode = True Then
		Response.Write "(units) <input type='text' class='text' style='width:60%' id='regulator_replaced_range_units' name='regulator_replaced_range_units' value='" & regulator_replaced_range_units & "' onchange='setupdate();' /></td>"
	Else
		Response.Write regulator_replaced_range_units & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Installed By: </td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='installed_by' name='installed_by' value='" & installed_by & "' onchange='setupdate();' /></td>"
	Else
		Response.Write installed_by & "</td>"
	End If
	Response.Write "<td id='formtd'>Date: </td>"
	Response.Write "<td id='formtd' colspan='4'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' id='installed_date' name='installed_date' value='" & installed_date & "' onchange='chkDate(this);setupdate();' />"
		Response.Write "<a href='javascript: void(0);' onclick='displayDatePicker(""installed_date"");setupdate();return false;'><img src='../images/calendar.bmp' alt='Calendar' style='vertical-align:top'></a></td>"
	Else
		Response.Write installed_date & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='7'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='7'>COMMENT:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='7'>"
	If editMode = True Then
		Response.Write "<textarea style='width:100%;height:40px' rows='2' cols='80' id='comment' name='comment' onchange='setupdate();'>" & comment & "</textarea></td>"
	Else
		Response.Write comment & "</td>"
	End If
	Response.Write "</tr>"

	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='7'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='7'>DISCREPANCY:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='7' style='font-weight:bold'>&nbsp;&nbsp;COMMENTS:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='7'>"
	If editMode = True Then
		Response.Write "<textarea style='width:100%;height:40px' rows='2' cols='80' id='discrepency_comments' name='discrepency_comments' onchange='setupdate();'>" & discrepency_comments & "</textarea></td>"
	Else
		Response.Write discrepency_comments & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='7' style='font-weight:bold'>&nbsp;&nbsp;FOLLOW-UP:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='7'>"
	If editMode = True Then
		Response.Write "<textarea style='width:100%;height:40px' rows='2' cols='80' id='discrepency_followup' name='discrepency_followup' onchange='setupdate();'>" & discrepency_followup & "</textarea></td>"
	Else
		Response.Write discrepency_followup & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "</table>"

	If editMode = True Then
		Response.Write "<br />"
		Response.Write "<div style='text-align:center'>"
		Response.Write "<button type='button' class='noprint' id='submit1' name='submit1' onclick='saveData();'>Submit</button>"
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
