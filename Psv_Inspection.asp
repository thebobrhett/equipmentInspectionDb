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
 window.open("http://mogsb8/inspections/psv_technicaldata.asp?itemID="+document.form1.itemID.value+"&edit=false","TechnicalData");
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
' Keith Brooks - Monday, January 17, 2011
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
'Inspection items
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
	
	'Create the update flag that can be set when a field is changed to allow the
	'user to be reminded to submit the form.
'	Response.Write "<input type='hidden' id='updateFlag' name='updateFlag' value='false' />"
	
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
				selected_orifice_size = rs2("selected_orifice_size")
				selected_orifice_size_units = rs2("selected_orifice_size_units")
				minimum_capacity = rs2("minimum_capacity")
				minimum_capacity_units = rs2("minimum_capacity_units")
				out_connect_size = rs2("out_connect_size")
				out_connect_size_units = rs2("out_connect_size_units")
				in_connect_size = rs2("in_connect_size")
				in_connect_size_units = rs2("in_connect_size_units")
				set_pressure = rs2("set_pressure")
				set_pressure_units = rs2("set_pressure_units")
				name_plate_stamp = rs2("name_plate_stamp")
				fluid_state = rs2("fluid_state")
				fluid_service = rs2("fluid_service")
				specification_number = rs2("specification_number")
				serial_number = rs2("serial_number")
				body_material = rs2("body_material")
				trim_material = rs2("trim_material")
				noz_disk_material = rs2("noz_disk_material")
				spring_material = rs2("spring_material")
				gasket_material = rs2("gasket_material")
				bellows_material = rs2("bellows_material")
			End If
			rs2.Close
			Set rs2 = Nothing
			
			'Fill in the existing inspection data variables.
			If IsNull(rs("inspection_date")) Then
				inspection_date = ""
			Else
				inspection_date = FormatDateTime(rs("inspection_date"),2)
			End If
			valve_nameplate_matches = rs("valve_nameplate_matches")
			valve_decontaminated = rs("valve_decontaminated")
			inlet_piping_condition = rs("inlet_piping_condition")
			inlet_piping_required_cleaning = rs("inlet_piping_required_cleaning")
			outlet_piping_condition = rs("outlet_piping_condition")
			outlet_piping_required_cleaning = rs("outlet_piping_required_cleaning")
			field_initial_inspection_other = rs("field_initial_inspection_other")
			condition_of_body = rs("condition_of_body")
			inspection_company = rs("inspection_company")
			inspected_by = rs("inspected_by")
			If IsNull(rs("inspected_date")) Then
				inspected_date = ""
			Else
				inspected_date = FormatDateTime(rs("inspected_date"),2)
			End If
			leaked_at_90pct_of_test_pr = rs("leaked_at_90pct_of_test_pr")
			popped = rs("popped")
			operated_properly = rs("operated_properly")
			returned_to_service = rs("returned_to_service")
			test_conducted_by = rs("test_conducted_by")
			If IsNull(rs("test_conducted_date")) Then
				test_conducted_date = ""
			Else
				test_conducted_date = FormatDateTime(rs("test_conducted_date"),2)
			End If
			test_only = rs("test_only")
			overhaul = rs("overhaul")
			test_and_reset = rs("test_and_reset")
			scrap_and_replace = rs("scrap_and_replace")
			work_required_other = rs("work_required_other")
			disassembly_condition = rs("disassembly_condition")
			seats_lapped = rs("seats_lapped")
			guide_replaced = rs("guide_replaced")
			seats_replaced = rs("seats_replaced")
			spring_replaced = rs("spring_replaced")
			other_replaced = rs("other_replaced")
			other_repairs = rs("other_repairs")
			repair_company = rs("repair_company")
			repaired_by = rs("repaired_by")
			If IsNull(rs("repaired_date")) Then
				repaired_date = ""
			Else
				repaired_date = FormatDateTime(rs("repaired_date"),2)
			End If
			code_stamp = rs("code_stamp")
			authorization_number = rs("authorization_number")
			work_order_number = rs("work_order_number")
			next_inspection_due = rs("next_inspection_due")
			previous_inspection = rs("previous_inspection")
			set_frequency = rs("set_frequency")
			set_frequency_units = rs("set_frequency_units")
			policy_insp_verify = rs("policy_insp_verify")
			repair_required = rs("repair_required")
			repair_type = rs("repair_type")
			bubble_tight_at = rs("bubble_tight_at")
			final_test_type = rs("final_test_type")
			final_test_set_press = rs("final_test_set_press")
			final_test_set_temp = rs("final_test_set_temp")
			final_test_by = rs("final_test_by")
			final_test_date = rs("final_test_date")
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
			
			'Initialize the inspection data variables.
			inspection_date = Date
			valve_nameplate_matches = 0
			valve_decontaminated = 0
			inlet_piping_condition = ""
			inlet_piping_required_cleaning = 0
			outlet_piping_condition = ""
			outlet_piping_required_cleaning = 0
			field_initial_inspection_other = ""
			condition_of_body = ""
			inspection_company = ""
			inspected_by = ""
			inspected_date = Date
			leaked_at_90pct_of_test_pr = 0
			popped = 0
			operated_properly = 0
			returned_to_service = 0
			test_conducted_by = ""
			test_conducted_date = ""
			test_only = 0
			overhaul = 0
			test_and_reset = 0
			scrap_and_replace = 0
			work_required_other = ""
			disassembly_condition = ""
			seats_lapped = 0
			guide_replaced = 0
			seats_replaced = 0
			spring_replaced = 0
			other_replaced = ""
			other_repairs = ""
			repair_company = ""
			repaired_by = ""
			repaired_date = ""
			code_stamp = ""
			authorization_number = ""
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
			bubble_tight_at = ""
			final_test_type = ""
			final_test_set_press = ""
			final_test_set_temp = ""
			final_test_by = ""
			final_test_date = ""
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
	Response.Write "<td id='grouptd' colspan='5'>RELIEF VALVE DESIGN:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='techtd' style='width:25%'>Manufacturer: </td>"
	Response.Write "<td id='techtd' colspan='2'>" & manufacturer & "</td>"
	Response.Write "<td id='techtd' style='width:20%'>Spec Number: </td>"
	Response.Write "<td id='techtd' style='width:30%'>" & specification_number & "</td>"
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
	Response.Write "<td id='techtd' style='width:15%'>" & selected_orifice_size_units & "</td>"
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

	'Draw the button to open the technical data form.
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='5' style='text-align:right'>"
	Response.Write "<input type='button' class='noprint' id='techdata' name='techdata' value='Technical Data' onclick='opentechdata();return false;' /></td>"
	Response.Write "</tr>"
	
	'Draw the data entry section.
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='5'>FIELD/INITIAL INSPECTION:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:37%'>Valve Nameplate Data Matches Above Data: </td>"
	Response.Write "<td id='formtd' style='width:13%'>"
	If valve_nameplate_matches = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='valve_nameplate_matches' name='valve_nameplate_matches' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='valve_nameplate_matches' name='valve_nameplate_matches' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd' style='width:20%'>Valve Decontaminated: </td>"
	Response.Write "<td id='formtd' style='width:30%'>"
	If valve_decontaminated = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='valve_decontaminated' name='valve_decontaminated' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='valve_decontaminated' name='valve_decontaminated' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='2' style='text-align:center'>INLET PIPING</td>"
	Response.Write "<td id='grouptd' colspan='2' style='text-align:center'>OUTLET PIPING</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:25%'>Condition: </td>"
	Response.Write "<td id='formtd' style='width:25%'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='inlet_piping_condition' name='inlet_piping_condition' value='" & inlet_piping_condition & "' onchange='setupdate();' /></td>"
	Else
		Response.Write inlet_piping_condition & "</td>"
	End If
	Response.Write "<td id='formtd' style='width:20%'>Condition: </td>"
	Response.Write "<td id='formtd' style='width:30%'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='outlet_piping_condition' name='outlet_piping_condition' value='" & outlet_piping_condition & "' onchange='setupdate();' /></td>"
	Else
		Response.Write outlet_piping_condition & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Required Cleaning: </td>"
	Response.Write "<td id='formtd'>"
	If inlet_piping_required_cleaning = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='inlet_piping_required_cleaning' name='inlet_piping_required_cleaning' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='inlet_piping_required_cleaning' name='inlet_piping_required_cleaning' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Required Cleaning: </td>"
	Response.Write "<td id='formtd'>"
	If outlet_piping_required_cleaning = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='outlet_piping_required_cleaning' name='outlet_piping_required_cleaning' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='outlet_piping_required_cleaning' name='outlet_piping_required_cleaning' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='vertical-align:top'>Other (Specify): </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	If editMode = True Then
		Response.Write "<textarea style='width:100%;height:40px' id='field_initial_inspection_other' name='field_initial_inspection_other' rows='2' cols='80' onchange='setupdate();'>" & field_initial_inspection_other & "</textarea></td>"
	Else
		Response.Write field_initial_inspection_other & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Condition of Body (Specify): </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='condition_of_body' name='condition_of_body' value='" & condition_of_body & "' onchange='setupdate();' /></td>"
	Else
		Response.Write condition_of_body & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr><td id='formtd' colspan='4'>&nbsp;</td></tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='font-weight:bold'>Inspection Company: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='inspection_company' name='inspection_company' value='" & inspection_company & "' onchange='setupdate();' /></td>"
	Else
		Response.Write inspection_company & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='font-weight:bold'>Inspected By: </td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='inspected_by' name='inspected_by' value='" & inspected_by & "' onchange='setupdate();' /></td>"
	Else
		Response.Write inspected_by & "</td>"
	End If
	Response.Write "<td id='formtd' style='font-weight:bold'>Date: </td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' id='inspected_date' name='inspected_date' value='" & inspected_date & "' onchange='chkDate(this);setupdate();' />"
		Response.Write "<a href='javascript: void(0);' onclick='displayDatePicker(""inspected_date"");setupdate();return false;'><img src='../images/calendar.bmp' alt='Calendar' style='vertical-align:top'></a></td>"
	Else
		Response.Write inspected_date & "</td>"
	End If
	Response.Write "</tr>"

	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='4'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='4'>INITIAL TEST:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Leaked at 90% of Test Pr.: </td>"
	Response.Write "<td id='formtd'>"
	If leaked_at_90pct_of_test_pr = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='leaked_at_90pct_of_test_pr' name='leaked_at_90pct_of_test_pr' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='leaked_at_90pct_of_test_pr' name='leaked_at_90pct_of_test_pr' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Popped: </td>"
	Response.Write "<td id='formtd'>"
	If popped = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='popped' name='popped' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='popped' name='popped' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Operated Properly: </td>"
	Response.Write "<td id='formtd'>"
	If operated_properly = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='operated_properly' name='operated_properly' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='operated_properly' name='operated_properly' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Returned To Service: </td>"
	Response.Write "<td id='formtd'>"
	If returned_to_service = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='returned_to_service' name='returned_to_service' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='returned_to_service' name='returned_to_service' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Test Conducted By: </td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='test_conducted_by' name='test_conducted_by' value='" & test_conducted_by & "' onchange='setupdate();' /></td>"
	Else
		Response.Write test_conducted_by & "</td>"
	End If
	Response.Write "<td id='formtd'>Date: </td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' id='test_conducted_date' name='test_conducted_date' value='" & test_conducted_date & "' onchange='chkDate(this);setupdate();' />"
		Response.Write "<a href='javascript: void(0);' onclick='displayDatePicker(""test_conducted_date"");setupdate();return false;'><img src='../images/calendar.bmp' alt='Calendar' style='vertical-align:top'></a></td>"
	Else
		Response.Write test_conducted_date & "</td>"
	End If
	Response.Write "</tr>"

	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='4'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='4'>WORK REQUIRED:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Test Only: </td>"
	Response.Write "<td id='formtd'>"
	If test_only = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='test_only' name='test_only' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='test_only' name='test_only' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Overhaul: </td>"
	Response.Write "<td id='formtd'>"
	If overhaul = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='overhaul' name='overhaul' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='overhaul' name='overhaul' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Test and Reset: </td>"
	Response.Write "<td id='formtd'>"
	If test_and_reset = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='test_and_reset' name='test_and_reset' " & field_disabled & " value='1' onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='test_and_reset' name='test_and_reset' " & field_disabled & " value='1' onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Scrap and Replace: </td>"
	Response.Write "<td id='formtd'>"
	If scrap_and_replace = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='scrap_and_replace' name='scrap_and_replace' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='scrap_and_replace' name='scrap_and_replace' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Other (Specify): </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='work_required_other' name='work_required_other' value='" & work_required_other & "' onchange='setupdate();' /></td>"
	Else
		Response.Write work_required_other & "</td>"
	End If
	Response.Write "</tr>"

	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='4'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='4'>REPAIRS MADE:</td>"
	Response.Write "</tr>"
	Response.Write "<td id='formtd'>Disassembly Condition: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='disassembly_condition' name='disassembly_condition' value='" & disassembly_condition & "' onchange='setupdate();' /></td>"
	Else
		Response.Write disassembly_condition & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Seats Lapped: </td>"
	Response.Write "<td id='formtd'>"
	If seats_lapped = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='seats_lapped' name='seats_lapped' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='seats_lapped' name='seats_lapped' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Guide Replaced: </td>"
	Response.Write "<td id='formtd'>"
	If guide_replaced = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='guide_replaced' name='guide_replaced' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='guide_replaced' name='guide_replaced' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "</tr>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Seats Replaced: </td>"
	Response.Write "<td id='formtd'>"
	If seats_replaced = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='seats_replaced' name='seats_replaced' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='seats_replaced' name='seats_replaced' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "<td id='formtd'>Spring Replaced: </td>"
	Response.Write "<td id='formtd'>"
	If spring_replaced = 0 Then
		Response.Write "<input type='checkbox' class='checkbox' id='spring_replaced' name='spring_replaced' value='1' " & field_disabled & " onchange='setupdate();' /></td>"
	Else
		Response.Write "<input type='checkbox' class='checkbox' id='spring_replaced' name='spring_replaced' value='1' " & field_disabled & " onchange='setupdate();' checked /></td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Other Replaced (Specify): </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='other_replaced' name='other_replaced' value='" & other_replaced & "' onchange='setupdate();' /></td>"
	Else
		Response.Write other_replaced & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Other Repairs: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='other_repairs' name='other_repairs' value='" & other_repairs & "' onchange='setupdate();' /></td>"
	Else
		Response.Write other_repairs & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Repair Company: </td>"
	Response.Write "<td id='formtd' colspan='3'>"
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
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' id='repaired_date' name='repaired_date' value='" & repaired_date & "' onchange='chkDate(this);setupdate();' />"
		Response.Write "<a href='javascript: void(0);' onclick='displayDatePicker(""repaired_date"");setupdate();return false;'><img src='../images/calendar.bmp' alt='Calendar' style='vertical-align:top'></a></td>"
	Else
		Response.Write repaired_date & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Code Stamp: </td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='code_stamp' name='code_stamp' value='" & code_stamp & "' onchange='setupdate();' /></td>"
	Else
		Response.Write code_stamp & "</td>"
	End If
	Response.Write "<td id='formtd'>Authorization No: </td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='authorization_number' name='authorization_number' value='" & authorization_number & "' onchange='setupdate();' /></td>"
	Else
		Response.Write authorization_number & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='2' style='padding:0px'>"
	Response.Write "<table style='width:100%;border:1px solid black'>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:40%'>Policy Insp. Verify:</td>"
	Response.Write "<td id='formtd' style='width:20%'>"
	If LCase(policy_insp_verify) = "yes" Then
		Response.Write "<input type='radio' class='radio' id='policy_insp_verify' name='policy_insp_verify' value='yes' " & field_disabled & " onchange='setupdate();' checked />yes</td>"
	Else
		Response.Write "<input type='radio' class='radio' id='policy_insp_verify' name='policy_insp_verify' value='yes' " & field_disabled & " onchange='setupdate();' />yes</td>"
	End If
	Response.Write "<td id='formtd' style='width:20%'>"
	If LCase(policy_insp_verify) = "no" Then
		Response.Write "<input type='radio' class='radio' id='policy_insp_verify' name='policy_insp_verify' value='no' " & field_disabled & " onchange='setupdate();' checked />no</td>"
	Else
		Response.Write "<input type='radio' class='radio' id='policy_insp_verify' name='policy_insp_verify' value='no' " & field_disabled & " onchange='setupdate();' />no</td>"
	End If
	Response.Write "<td id='formtd' style='width:20%'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Repair Required:</td>"
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
	Response.Write "<td id='formtd'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Repair Type:</td>"
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
	
	Response.Write "<div style='page-break-before:always;font-size:1;margin:0;border:0'><span style='visibility:hidden'>-</span></div>"
	
	Response.Write "<table style='width:100%;borders:none'>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='7'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='7'>FINAL TEST:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd' style='width:25%'>Bubble Tight At:</td>"
	Response.Write "<td id='formtd' style='width:7%'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='bubble_tight_at' name='bubble_tight_at' value='" & bubble_tight_at & "' onchange='chkNumeric(this);setupdate();' /></td>"
	Else
		Response.Write bubble_tight_at & "</td>"
	End If
	Response.Write "<td id='formtd' colspan='3'>PSIG</td>"
	Response.Write "<td id='formtd' style='width:20%'>Type:</td>"
	Response.Write "<td id='formtd' style='width:30%'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='final_test_type' name='final_test_type' value='" & final_test_type & "' onchange='setupdate();' /></td>"
	Else
		Response.Write final_test_type & "</td>"
	End If
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Set: </td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='final_test_set_press' name='final_test_set_press' value='" & final_test_set_press & "' onchange='chkNumeric(this);setupdate();' /></td>"
	Else
		Response.Write final_test_set_press & "</td>"
	End If
	Response.Write "<td id='formtd' style='width:8%'>PSIG at</td>"
	Response.Write "<td id='formtd' style='width:7%'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:100%' id='final_test_set_temp' name='final_test_set_temp' value='" & final_test_set_temp & "' onchange='chkNumeric(this);setupdate();' /></td>"
	Else
		Response.Write final_test_set_temp & "</td>"
	End If
	Response.Write "<td id='formtd' style='width:2%'>F</td>"
	Response.Write "<td id='formtd' colspan='2'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Final Test By: </td>"
	Response.Write "<td id='formtd' colspan='4'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='final_test_by' name='final_test_by' value='" & final_test_by & "' onchange='setupdate();' /></td>"
	Else
		Response.Write final_test_by & "</td>"
	End If
	Response.Write "<td id='formtd'>Date: </td>"
	Response.Write "<td id='formtd'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' id='final_test_date' name='final_test_date' value='" & final_test_date & "' onchange='chkDate(this);setupdate();' />"
		Response.Write "<a href='javascript: void(0);' onclick='displayDatePicker(""final_test_date"");setupdate();return false;'><img src='../images/calendar.bmp' alt='Calendar' style='vertical-align:top'></a></td>"
	Else
		Response.Write final_test_date & "</td>"
	End If
	Response.Write "</tr>"

	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='7'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='grouptd' colspan='7'>INSTALLATION:</td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td id='formtd'>Installed By: </td>"
	Response.Write "<td id='formtd' colspan='4'>"
	If editMode = True Then
		Response.Write "<input type='text' class='text' style='width:90%' id='installed_by' name='installed_by' value='" & installed_by & "' onchange='setupdate();' /></td>"
	Else
		Response.Write installed_by & "</td>"
	End If
	Response.Write "<td id='formtd'>Date: </td>"
	Response.Write "<td id='formtd'>"
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
	Response.Write "<tr>"
	Response.Write "<td id='formtd' colspan='7'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	
	If editMode = True Then
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
