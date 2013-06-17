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
 document.form1.next_inspection_date.value=t.getMonth()+1+"/"+t.getDate()+"/"+t.getFullYear();
 needToConfirm=true;
}
function canceledit() {
 //Put the page back in read-only mode.
 window.location.href='psv_technicaldata.asp?itemid='+document.form1.itemID.value+'&edit=false';
}
function editmode() {
 //Put the page in edit mode.
 window.location.href='psv_technicaldata.asp?itemid='+document.form1.itemID.value+'&edit=true';
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
<title>PSV Technical Data</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="InspectionsFunctions.asp"-->
<link rel=STYLESHEET href='formstyle.css' type='text/css'>
</head>
<%
'*************
' Revision History
' 
' Keith Brooks - Tuesday, February 8, 2011
'   Creation.
'*************

Dim currentuser
Dim access
Dim editMode
Dim field_disabled
Dim cn
Dim rs
Dim rs2
Dim sqlString
Dim equipment_item_tag
Dim assembly
Dim equipment_item_name
Dim area
Dim technical_data_id
Dim asset_type
Dim installed_date
Dim file_complete
Dim report_on_file
Dim inspection_frequency
Dim inspection_frequency_units
Dim next_inspection_date
Dim previous_inspection_date
Dim inspection_scheduled
Dim status_code
Dim shutdown_required
Dim job_plan_type
Dim deviation_sequence
Dim deviation_approval
Dim equip_sequence_no
Dim manufacturer
Dim model_number
Dim serial_number
Dim specification_number
Dim revision_code
Dim valve_type
Dim fluid_service
Dim fluid_state
Dim calculated_orifice_size
Dim calculated_orifice_size_units
Dim selected_orifice_size
Dim selected_orifice_size_units
Dim orifice_designation
Dim in_connect_size
Dim in_connect_size_units
Dim in_connect_type
Dim out_connect_size
Dim out_connect_size_units
Dim out_connect_type
Dim bonnet_type
Dim cap_type
Dim lifting_gear_type
Dim gag
Dim name_plate_stamp
Dim body_material
Dim bonnet_material
Dim trim_material
Dim noz_disk_material
Dim bellows_material
Dim spring_material
Dim gasket_material
Dim set_pressure
Dim set_pressure_units
Dim set_temperature
Dim set_temperature_units
Dim accumulation_pct
Dim blowdown_pct
Dim minimum_capacity
Dim minimum_capacity_units
Dim minimum_capacity_code
Dim normal_temp
Dim normal_temp_units
Dim max_accum_temp
Dim max_accum_temp_units
Dim specific_gravity
Dim viscosity
Dim viscosity_units
Dim molecular_weight
Dim compress_factor
Dim normal_press
Dim normal_press_units
Dim super_back_press
Dim super_back_press_units
Dim diff_set_press
Dim diff_set_press_units
Dim builtup_back_press
Dim builtup_back_press_units
Dim discharge_to
Dim used_in_combo
Dim comb_dev_tag_no
Dim comb_dev_mfgr
Dim comb_dev_model
Dim comb_dev_size
Dim derating_factor
Dim marking_tag_required
Dim drawing_number
Dim inspection_standard

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("inspections", "psv_technicaldata", currentuser)
If access <> "none" Then

	If LCase(Request("edit")) = "true" And (access = "write" Or access = "delete") Then
		editMode = True
		field_disabled = ""
		Response.Write "<body  onload='document.form1.asset_type.focus();'>"
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
	set rs2 = CreateObject("adodb.recordset")

	Response.Write "<form id='form1' name='form1' action='technicaldataaction.asp' method='post'>"
	
	'Save the equipment type for use by the action page.
	Response.Write "<input type='hidden' id='equipType' name='equipType' value='psv' />"
	Response.Write "<input type='hidden' id='equipTypeID' name='equipTypeID' value='" & Request("equipTypeID") & "' />"

	'Save the plant area for use by the action page.
	Response.Write "<input type='hidden' id='plantarea' name='plantarea' value='" & Request("plantarea") & "' />"

	If Request("itemID") <> "" Then

		'Save the item ID.
		Response.Write "<input type='hidden' id='itemID' name='itemID' value='" & Request("itemID") & "' />"
		
		'Get the equipment data for the header.
		sqlString = "SELECT equipment_item_tag,equipment_item_name,assembly,area " & _
				"FROM equipment_items WHERE equipment_item_id=" & Request("itemID")
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			
			'Load the equipment data variables.
			equipment_item_tag = rs("equipment_item_tag")
			assembly = rs("assembly")
			equipment_item_name = rs("equipment_item_name")
			area = rs("area")
			rs.Close

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
			Response.Write "<td style='text-align:left;vertical-align:top;font-size:12pt;font-weight:bold'>Relief Valve - Technical Data Sheet</td>"
			Response.Write "<td style='text-align:right;font-size:14pt;font-weight:bold'>AKSA</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td colspan='2' style='border-top:1px solid black;border-bottom:1px solid black;height:2px;padding:0px;line-height:2px'>&nbsp;</td>"
			Response.Write "</tr>"
			Response.Write "</table>"
			'Draw the form.
			Response.Write "<table style='width:100%;border:none'>"
			Response.Write "<tr>"
			Response.Write "<td id='formtd' style='width:20%'>Equip No: </td>"
			Response.Write "<td id='formtd' style='width:30%'>" & equipment_item_tag & "</td>"
			Response.Write "<td id='formtd' style='width:20%'>Assembly: </td>"
			Response.Write "<td id='formtd' style='width:30%'>" & assembly & "</td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='formtd'>Equip Name: </td>"
			Response.Write "<td id='formtd'>" & equipment_item_name & "</td>"
			Response.Write "<td id='formtd'>Area: </td>"
			Response.Write "<td id='formtd'>" & area & "</td>"
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
	
			'Get the technical data.
			sqlString = "SELECT * FROM psv_technical_data " & _
					"WHERE equipment_item_id=" & Request("itemID")
			Set rs = cn.Execute(sqlString)
			If Not rs.BOF Then
				'If the record exists, assign the existing values to the variables.
				rs.MoveFirst
				technical_data_id = rs("technical_data_id")
				asset_type = rs("asset_type")
				installed_date = rs("installed_date")
				file_complete = rs("file_complete")
				report_on_file = rs("report_on_file")
				If Not IsNull(rs("inspection_frequency")) Then
					inspection_frequency = rs("inspection_frequency")
				Else
					inspection_frequency = 0
				End If
				If Not IsNull(rs("inspection_frequency_units")) Then
					inspection_frequency_units = rs("inspection_frequency_units")
				Else
					inspection_frequency_units = ""
				End If
				next_inspection_date = rs("next_inspection_date")
				inspection_scheduled = rs("inspection_scheduled")
				status_code = rs("status_code")
				shutdown_required = rs("shutdown_required")
				job_plan_type = rs("job_plan_type")
				deviation_sequence = rs("deviation_sequence")
				deviation_approval = rs("deviation_approval")
				equip_sequence_no = rs("equip_sequence_no")
				manufacturer = rs("manufacturer")
				model_number = rs("model_number")
				serial_number = rs("serial_number")
				specification_number = rs("specification_number")
				revision_code = rs("revision_code")
				valve_type = rs("valve_type")
				fluid_service = rs("fluid_service")
				fluid_state = rs("fluid_state")
				calculated_orifice_size = rs("calculated_orifice_size")
				calculated_orifice_size_units = rs("calculated_orifice_size_units")
				selected_orifice_size = rs("selected_orifice_size")
				selected_orifice_size_units = rs("selected_orifice_size_units")
				orifice_designation = rs("orifice_designation")
				in_connect_size = rs("in_connect_size")
				in_connect_size_units = rs("in_connect_size_units")
				in_connect_type = rs("in_connect_type")
				out_connect_size = rs("out_connect_size")
				out_connect_size_units = rs("out_connect_size_units")
				out_connect_type = rs("out_connect_type")
				bonnet_type = rs("bonnet_type")
				cap_type = rs("cap_type")
				lifting_gear_type = rs("lifting_gear_type")
				gag = rs("gag")
				name_plate_stamp = rs("name_plate_stamp")
				body_material = rs("body_material")
				bonnet_material = rs("bonnet_material")
				trim_material = rs("trim_material")
				noz_disk_material = rs("noz_disk_material")
				bellows_material = rs("bellows_material")
				spring_material = rs("spring_material")
				gasket_material = rs("gasket_material")
				set_pressure = rs("set_pressure")
				set_pressure_units = rs("set_pressure_units")
				set_temperature = rs("set_temperature")
				set_temperature_units = rs("set_temperature_units")
				accumulation_pct = rs("accumulation_pct")
				blowdown_pct = rs("blowdown_pct")
				minimum_capacity = rs("minimum_capacity")
				minimum_capacity_units = rs("minimum_capacity_units")
				minimum_capacity_code = rs("minimum_capacity_code")
				normal_temp = rs("normal_temp")
				normal_temp_units = rs("normal_temp_units")
				max_accum_temp = rs("max_accum_temp")
				max_accum_temp_units = rs("max_accum_temp_units")
				specific_gravity = rs("specific_gravity")
				viscosity = rs("viscosity")
				viscosity_units = rs("viscosity_units")
				molecular_weight = rs("molecular_weight")
				compress_factor = rs("compress_factor")
				normal_press = rs("normal_press")
				normal_press_units = rs("normal_press_units")
				super_back_press = rs("super_back_press")
				super_back_press_units = rs("super_back_press_units")
				diff_set_press = rs("diff_set_press")
				diff_set_press_units = rs("diff_set_press_units")
				builtup_back_press = rs("builtup_back_press")
				builtup_back_press_units = rs("builtup_back_press_units")
				discharge_to = rs("discharge_to")
				used_in_combo = rs("used_in_combo")
				comb_dev_tag_no = rs("comb_dev_tag_no")
				comb_dev_mfgr = rs("comb_dev_mfgr")
				comb_dev_model = rs("comb_dev_model")
				comb_dev_size = rs("comb_dev_size")
				derating_factor = rs("derating_factor")
				marking_tag_required = rs("marking_tag_required")
				drawing_number = rs("drawing_number")
				inspection_standard = rs("inspection_standard")
				
			Else
				'If the record doesn't exist, initialize the variables.
				technical_data_id = 0
				asset_type = ""
				installed_date = ""
				file_complete = ""
				report_on_file = ""
				inspection_frequency = 0
				inspection_frequency_units = ""
				next_inspection_date = ""
				inspection_scheduled = ""
				status_code = ""
				shutdown_required = ""
				job_plan_type = ""
				deviation_sequence = ""
				deviation_approval = ""
				equip_sequence_no = ""
				manufacturer = ""
				model_number = ""
				serial_number = ""
				specification_number = ""
				revision_code = ""
				valve_type = ""
				fluid_service = ""
				fluid_state = ""
				calculated_orifice_size = ""
				calculated_orifice_size_units = ""
				selected_orifice_size = ""
				selected_orifice_size_units = ""
				orifice_designation = ""
				in_connect_size = ""
				in_connect_size_units = ""
				in_connect_type = ""
				out_connect_size = ""
				out_connect_size_units = ""
				out_connect_type = ""
				bonnet_type = ""
				cap_type = ""
				lifting_gear_type = ""
				gag = ""
				name_plate_stamp = ""
				body_material = ""
				bonnet_material = ""
				trim_material = ""
				noz_disk_material = ""
				bellows_material = ""
				spring_material = ""
				gasket_material = ""
				set_pressure = ""
				set_pressure_units = ""
				set_temperature = ""
				set_temperature_units = ""
				accumulation_pct = ""
				blowdown_pct = ""
				minimum_capacity = ""
				minimum_capacity_units = ""
				minimum_capacity_code = ""
				normal_temp = ""
				normal_temp_units = ""
				max_accum_temp = ""
				max_accum_temp_units = ""
				specific_gravity = ""
				viscosity = ""
				viscosity_units = ""
				molecular_weight = ""
				compress_factor = ""
				normal_press = ""
				normal_press_units = ""
				super_back_press = ""
				super_back_press_units = ""
				diff_set_press = ""
				diff_set_press_units = ""
				builtup_back_press = ""
				builtup_back_press_units = ""
				discharge_to = ""
				used_in_combo = ""
				comb_dev_tag_no = ""
				comb_dev_mfgr = ""
				comb_dev_model = ""
				comb_dev_size = ""
				derating_factor = ""
				marking_tag_required = ""
				drawing_number = ""
				inspection_standard = ""
			
			End If
			
			'If inspection frequency or its units are not specified, look the default values
			'up in the equipment types table.
			If inspection_frequency = 0 Or inspection_frequency_units = "" Then
				sqlString = "SELECT inspection_interval,inspection_interval_units " & _
						"FROM equipment_types WHERE equipment_type_name='PSV'"
				Set rs2 = cn.Execute(sqlString)
				If Not rs2.BOF Then
					rs2.MoveFirst
					If Not IsNull(rs2("inspection_interval")) And inspection_frequency = 0 Then
						inspection_frequency = rs2("inspection_interval")
					End If
					If Not IsNull(rs2("inspection_interval_units")) And inspection_frequency_units = "" Then
						inspection_frequency_units = rs2("inspection_interval_units")
					End If
				End If
				rs2.Close
			End If
			
			'Get the previous inspection date, if it exists.
			sqlString = "SELECT MAX(inspection_date) " & _
					"FROM psv_inspection_data " & _
					"WHERE equipment_item_id=" & Request("itemID")
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				If Not IsNull(rs2(0)) Then
					previous_inspection_date = FormatDateTime(rs2(0),2)
				Else
					previous_inspection_date = Date
				End If
			Else
				previous_inspection_date = Date
			End If
			rs2.Close

			'Calculate the next inspection date if it is not specified.
			If IsNull(next_inspection_date) Or next_inspection_date = "" Then
				If CInt(inspection_frequency) > 0 And inspection_frequency_units <> "" Then
					Dim units
					Select Case LCase(inspection_frequency_units)
						Case "years","year","yyyy","yy","y"
							units = "yyyy"
						Case "months","month","mon","m"
							units = "m"
						Case "days","day","dd","d"
							units = "d"
						Case Else
							units = ""
					End Select
					next_inspection_date = DateAdd(units,inspection_frequency,previous_inspection_date)
				Else
					next_inspection_date = ""
				End If
			End If

			'Draw the technical data body.
			Response.Write "<tbody><tr><td>"
			Response.Write "<input type='hidden' id='technical_data_id' name='technical_data_id' value='" & technical_data_id & "' />"
			Response.Write "<table style='width:100%;border:none'>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd' style='width:20%'>Asset Type:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='asset_type' name='asset_type' value='" & asset_type & "' onchange='setupdate();' /></td>"
			Else
				Response.Write asset_type & "</td>"
			End If
			Response.Write "<td id='techtd' style='width:20%'>Name Plate Stamp:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='name_plate_stamp' name='name_plate_stamp' value='" & name_plate_stamp & "' onchange='setupdate();' /></td>"
			Else
				Response.Write name_plate_stamp & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Installed Date:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' id='installed_date' name='installed_date' value='" & installed_date & "' onchange='chkDate(this);setupdate();' />"
				Response.Write "<a href='javascript: void(0);' onclick='displayDatePicker(""installed_date"");setupdate();return false;'><img src='../images/calendar.bmp' alt='Calendar' style='vertical-align:top'></a></td>"
			Else
				Response.Write installed_date & "</td>"
			End If
			Response.Write "<td id='techtd'>Body Material:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='body_material' name='body_material' value='" & body_material & "' onchange='setupdate();' /></td>"
			Else
				Response.Write body_material & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>File Complete:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='file_complete' name='file_complete' value='" & file_complete & "' onchange='setupdate();' /></td>"
			Else
				Response.Write file_complete & "</td>"
			End If
			Response.Write "<td id='techtd'>Bonnet Material:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='bonnet_material' name='bonnet_material' value='" & bonnet_material & "' onchange='setupdate();' /></td>"
			Else
				Response.Write bonnet_material & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Report On File:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='report_on_file' name='report_on_file' value='" & report_on_file & "' onchange='setupdate();' /></td>"
			Else
				Response.Write report_on_file & "</td>"
			End If
			Response.Write "<td id='techtd'>Trim Material:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='trim_material' name='trim_material' value='" & trim_material & "' onchange='setupdate();' /></td>"
			Else
				Response.Write trim_material & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Inspection Freq:</td>"
			Response.Write "<td id='techtd' style='width:10%'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='inspection_frequency' name='inspection_frequency' value='" & inspection_frequency & "' onchange='chkNumeric(this);addDate(""" & previous_inspection_date & """,document.form1.inspection_frequency.value,document.form1.inspection_frequency_units.value);return false;' /></td>"
			Else
				Response.Write inspection_frequency & "</td>"
			End If
			Response.Write "<td id='techtd' style='width:20%'>"
			If editMode = True Then
				Response.Write "<select id='inspection_frequency_units' name='inspection_frequency_units' onchange='addDate(""" & previous_inspection_date & """,document.form1.inspection_frequency.value,document.form1.inspection_frequency_units.value);return false;'>"
				Response.Write "<option value=''>"
				If inspection_frequency_units = "days" Then
					Response.Write "<option value='days' selected>days"
				Else
					Response.Write "<option value='days'>days"
				End If
				If inspection_frequency_units = "months" Then
					Response.Write "<option value='months' selected>months"
				Else
					Response.Write "<option value='months'>months"
				End If
				If inspection_frequency_units = "years" Then
					Response.Write "<option value='years' selected>years"
				Else
					Response.Write "<option value='years'>years"
				End If
				Response.Write "</select></td>"
			Else
				Response.Write inspection_frequency_units & "</td>"
			End If
			Response.Write "<td id='techtd'>Noz Disk Material:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='noz_disk_material' name='noz_disk_material' value='" & noz_disk_material & "' onchange='setupdate();' /></td>"
			Else
				Response.Write noz_disk_material & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Next Inspect Date:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' id='next_inspection_date' name='next_inspection_date' value='" & next_inspection_date & "' onchange='chkDate(this);setupdate();' />"
				Response.Write "<a href='javascript: void(0);' onclick='displayDatePicker(""next_inspection_date"");setupdate();return false;'><img src='../images/calendar.bmp' alt='Calendar' style='vertical-align:top'></a></td>"
			Else
				Response.Write next_inspection_date & "</td>"
			End If
			Response.Write "<td id='techtd'>Bellows Material:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='bellows_material' name='bellows_material' value='" & bellows_material & "' onchange='setupdate();' /></td>"
			Else
				Response.Write bellows_material & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Inspection Scheduled:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='inspection_scheduled' name='inspection_scheduled' value='" & inspection_scheduled & "' onchange='setupdate();' /></td>"
			Else
				Response.Write inspection_scheduled & "</td>"
			End If
			Response.Write "<td id='techtd'>Spring Material:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='spring_material' name='spring_material' value='" & spring_material & "' onchange='setupdate();' /></td>"
			Else
				Response.Write spring_material & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Status Code:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<select id='status_code' name='status_code' onchange='setupdate();'>"
				Response.Write "<option value=''>"
				If UCase(status_code) = "ACTIVE" Then
					Response.Write "<option value='ACTIVE' selected>ACTIVE"
				Else
					Response.Write "<option value='ACTIVE'>ACTIVE"
				End If
				If UCase(status_code) = "SPARE" Then
					Response.Write "<option value='SPARE' selected>SPARE"
				Else
					Response.Write "<option value='SPARE'>SPARE"
				End If
				Response.Write "</select></td>"
			Else
				Response.Write status_code & "</td>"
			End If
			Response.Write "<td id='techtd'>Gasket Material:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='gasket_material' name='gasket_material' value='" & gasket_material & "' onchange='setupdate();' /></td>"
			Else
				Response.Write gasket_material & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Shutdown Required:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<select id='shutdown_required' name='shutdown_required' onchange='setupdate();'>"
				Response.Write "<option value=''>"
				If UCase(shutdown_required) = "YES" Then
					Response.Write "<option value='YES' selected>YES"
				Else
					Response.Write "<option value='YES'>YES"
				End If
				If UCase(shutdown_required) = "NO" Then
					Response.Write "<option value='NO' selected>NO"
				Else
					Response.Write "<option value='NO'>NO"
				End If
				Response.Write "</select></td>"
			Else
				Response.Write shutdown_required & "</td>"
			End If
			Response.Write "<td id='techtd'>Set Pressure:</td>"
			Response.Write "<td id='techtd' style='width:10%'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='set_pressure' name='set_pressure' value='" & set_pressure & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write set_pressure & "</td>"
			End If
			Response.Write "<td id='techtd' style='width:20%'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id=set_pressure_units' name='set_pressure_units' value='" & set_pressure_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write set_pressure_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Job Plan Type:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='job_plan_type' name='job_plan_type' value='" & job_plan_type & "' onchange='setupdate();' /></td>"
			Else
				Response.Write job_plan_type & "</td>"
			End If
			Response.Write "<td id='techtd'>Set Temperature:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='set_temperature' name='set_temperature' value='" & set_temperature & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write set_temperature & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='set_temperature_units' name='set_temperature_units' value='" & set_temperature_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write set_temperature_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Deviation Sequence:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='deviation_sequence' name='deviation_sequence' value='" & deviation_sequence & "' onchange='setupdate();' /></td>"
			Else
				Response.Write deviation_sequence & "</td>"
			End If
			Response.Write "<td id='techtd'>Accumulation Pct:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='accumulation_pct' name='accumulation_pct' value='" & accumulation_pct & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write accumulation_pct & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Deviation Approval:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='deviation_approval' name='deviation_approval' value='" & deviation_approval & "' onchange='setupdate();' /></td>"
			Else
				Response.Write deviation_approval & "</td>"
			End If
			Response.Write "<td id='techtd'>Blowdown Pct:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='blowdown_pct' name='blowdown_pct' value='" & blowdown_pct & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write blowdown_pct & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Equip Sequence No:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='equip_sequence_no' name='equip_sequence_no' value='" & equip_sequence_no & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write equip_sequence_no & "</td>"
			End If
			Response.Write "<td id='techtd'>Minimum Capacity:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='minimum_capacity' name='minimum_capacity' value='" & minimum_capacity & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write minimum_capacity & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='minimum_capacity_units' name='minimum_capacity_units' value='" & minimum_capacity_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write minimum_capacity_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Manufacturer:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='manufacturer' name='manufacturer' value='" & manufacturer & "' onchange='setupdate();' /></td>"
			Else
				Response.Write manufacturer & "</td>"
			End If
			Response.Write "<td id='techtd'>Min. Capacity Code:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='minimum_capacity_code' name='minimum_capacity_code' value='" & minimum_capacity_code & "' onchange='setupdate();' /></td>"
			Else
				Response.Write minimum_capacity_code & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Model No:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='model_number' name='model_number' value='" & model_number & "' onchange='setupdate();' /></td>"
			Else
				Response.Write model_number & "</td>"
			End If
			Response.Write "<td id='techtd'>Normal Temp:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='normal_temp' name='normal_temp' value='" & normal_temp & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write normal_temp & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='normal_temp_units' name='normal_temp_units' value='" & normal_temp_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write normal_temp_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Serial No:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='serial_number' name='serial_number' value='" & serial_number & "' onchange='setupdate();' /></td>"
			Else
				Response.Write serial_number & "</td>"
			End If
			Response.Write "<td id='techtd'>Max Accum Temp:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='max_accum_temp' name='max_accum_temp' value='" & max_accum_temp & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write max_accum_temp & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='max_accum_temp_units' name='max_accum_temp_units' value='" & max_accum_temp_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write max_accum_temp_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Specification No:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='specification_number' name='specification_number' value='" & specification_number & "' onchange='setupdate();' /></td>"
			Else
				Response.Write specification_number & "</td>"
			End If
			Response.Write "<td id='techtd'>Specific Gravity:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='specific_gravity' name='specific_gravity' value='" & specific_gravity & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write specific_gravity & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Revision Code:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='revision_code' name='revision_code' value='" & revision_code & "' onchange='setupdate();' /></td>"
			Else
				Response.Write revision_code & "</td>"
			End If
			Response.Write "<td id='techtd'>Viscosity:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='viscosity' name='viscosity' value='" & viscosity & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write viscosity & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='viscosity_units' name='viscosity_units' value='" & viscosity_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write viscosity_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Valve Type:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='valve_type' name='valve_type' value='" & valve_type & "' onchange='setupdate();' /></td>"
			Else
				Response.Write valve_type & "</td>"
			End If
			Response.Write "<td id='techtd'>Molecular Weight:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='molecular_weight' name='molecular_weight' value='" & molecular_weight & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write molecular_weight & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Fluid Service:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='fluid_service' name='fluid_service' value='" & fluid_service & "' onchange='setupdate();' /></td>"
			Else
				Response.Write fluid_service & "</td>"
			End If
			Response.Write "<td id='techtd'>Compress Factor:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='compress_factor' name='compress_factor' value='" & compress_factor & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write compress_factor & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Fluid State:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='fluid_state' name='fluid_state' value='" & fluid_state & "' onchange='setupdate();' /></td>"
			Else
				Response.Write fluid_state & "</td>"
			End If
			Response.Write "<td id='techtd'>Normal Press:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='normal_press' name='normal_press' value='" & normal_press & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write normal_press & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='normal_press_units' name='normal_press_units' value='" & normal_press_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write normal_press_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Calc Orifice:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='calculated_orifice_size' name='calculated_orifice_size' value='" & calculated_orifice_size & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write calculated_orifice_size & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='calculated_orifice_size_units' name='calculated_orifice_size_units' value='" & calculated_orifice_size_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write calculated_orifice_size_units & "</td>"
			End If
			Response.Write "<td id='techtd'>Super Back Press:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='super_back_press' name='super_back_press' value='" & super_back_press & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write super_back_press & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='super_back_press_units' name='super_back_press_units' value='" & super_back_press_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write super_back_press_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Select Orifice:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='selected_orifice_size' name='selected_orifice_size' value='" & selected_orifice_size & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write selected_orifice_size & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='selected_orifice_size_units' name='selected_orifice_size_units' value='" & selected_orifice_size_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write selected_orifice_size_units & "</td>"
			End If
			Response.Write "<td id='techtd'>Diff Set Press:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='diff_set_press' name='diff_set_press' value='" & diff_set_press & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write diff_set_press & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='diff_set_press_units' name='diff_set_press_units' value='" & diff_set_press_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write diff_set_press_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Orif Designation:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='orifice_designation' name='orifice_designation' value='" & orifice_designation & "' onchange='setupdate();' /></td>"
			Else
				Response.Write orifice_designation & "</td>"
			End If
			Response.Write "<td id='techtd'>Builtup Back Press:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='builtup_back_press' name='builtup_back_press' value='" & builtup_back_press & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write builtup_back_press & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='builtup_back_press_units' name='builtup_back_press_units' value='" & builtup_back_press_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write builtup_back_press_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>In Connect Size:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='in_connect_size' name='in_connect_size' value='" & in_connect_size & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write in_connect_size & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='in_connect_size_units' name='in_connect_size_units' value='" & in_connect_size_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write in_connect_size_units & "</td>"
			End If
			Response.Write "<td id='techtd'>Discharge To:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='discharge_to' name='discharge_to' value='" & discharge_to & "' onchange='setupdate();' /></td>"
			Else
				Response.Write discharge_to & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>In Connect Type:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='in_connect_type' name='in_connect_type' value='" & in_connect_type & "' onchange='setupdate();' /></td>"
			Else
				Response.Write in_connect_type & "</td>"
			End If
			Response.Write "<td id='techtd'>Used In Combo:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='used_in_combo' name='used_in_combo' value='" & used_in_combo & "' onchange='setupdate();' /></td>"
			Else
				Response.Write used_in_combo & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Out Connect Size:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='out_connect_size' name='out_connect_size' value='" & out_connect_size & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write out_connect_size & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='out_connect_size_units' name='out_connect_size_units' value='" & out_connect_size_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write out_connect_size_units & "</td>"
			End If
			Response.Write "<td id='techtd'>Comb Dev Tag No:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='comb_dev_tag_no' name='comb_dev_tag_no' value='" & comb_dev_tag_no & "' onchange='setupdate();' /></td>"
			Else
				Response.Write comb_dev_tag_no & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Out Connect Type:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='out_connect_type' name='out_connect_type' value='" & out_connect_type & "' onchange='setupdate();' /></td>"
			Else
				Response.Write out_connect_type & "</td>"
			End If
			Response.Write "<td id='techtd'>Comb Dev Mfgr:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='comb_dev_mfgr' name='comb_dev_mfgr' value='" & comb_dev_mfgr & "' onchange='setupdate();' /></td>"
			Else
				Response.Write comb_dev_mfgr & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Bonnet Type:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='bonnet_type' name='bonnet_type' value='" & bonnet_type & "' onchange='setupdate();' /></td>"
			Else
				Response.Write bonnet_type & "</td>"
			End If
			Response.Write "<td id='techtd'>Comb Dev Model:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='comb_dev_model' name='comb_dev_model' value='" & comb_dev_model & "' onchange='setupdate();' /></td>"
			Else
				Response.Write comb_dev_model & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Cap Type:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='cap_type' name='cap_type' value='" & cap_type & "' onchange='setupdate();' /></td>"
			Else
				Response.Write cap_type & "</td>"
			End If
			Response.Write "<td id='techtd'>Comb Dev Size:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='comb_dev_size' name='comb_dev_size' value='" & comb_dev_size & "' onchange='setupdate();' /></td>"
			Else
				Response.Write comb_dev_size & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Lifting Gear Type:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='lifting_gear_type' name='lifting_gear_type' value='" & lifting_gear_type & "' onchange='setupdate();' /></td>"
			Else
				Response.Write lifting_gear_type & "</td>"
			End If
			Response.Write "<td id='techtd'>Derating Factor:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='derating_factor' name='derating_factor' value='" & derating_factor & "' onchange='setupdate();' /></td>"
			Else
				Response.Write derating_factor & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>GAG:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='gag' name='gag' value='" & gag & "' onchange='setupdate();' /></td>"
			Else
				Response.Write gag & "</td>"
			End If
			Response.Write "<td id='techtd'>Marking Tag Required:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='marking_tag_required' name='marking_tag_required' value='" & marking_tag_required & "' onchange='setupdate();' /></td>"
			Else
				Response.Write marking_tag_required & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Drawing Number:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='drawing_number' name='drawing_number' value='" & drawing_number & "' onchange='setupdate();' /></td>"
			Else
				Response.Write drawing_number & "</td>"
			End If
			Response.Write "<td id='techtd'>Inspection Standard:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='inspection_standard' name='inspection_standard' value='" & inspection_standard & "' onchange='setupdate();' /></td>"
			Else
				Response.Write inspection_standard & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd colspan='6'>&nbsp;</td>"
			Response.Write "</tr>"
			If editMode = True Then
				Response.Write "<tr>"
				Response.Write "<td id='techtd' colspan='3' style='text-align:right'>"
				Response.Write "<button type='button' class='noprint' id='cancel' name='cancel' onclick='canceledit();return false;'>View</button></td>"
				Response.Write "<td id='techtd' colspan='3' style='text-align:left'>"
				Response.Write "<button type='button' class='noprint' id='submit1' name='submit1' onclick='saveData();'>Submit</button></td>"
				Response.Write "</tr>"
			ElseIf access = "write" Or access = "delete" Then
				Response.Write "<tr>"
				Response.Write "<td id='techtd' colspan='6' style='text-align:center'>"
				Response.Write "<input type='button' class='noprint' id='edit' name='edit' value='Edit' onclick='editmode();return false;' /></td>"
				Response.Write "</tr>"
			End If
			Response.Write "</table>"
			Response.Write "</td></tr></tbody>"	
			Response.Write "</table>"
	
		Else
			Response.Write "<h2>Item ID not found.</h2>"
		End If
		rs.Close
	Else
		Response.Write "<h2>Item ID not specified.</h2>"
	End If

	Set rs = Nothing
	Set rs2 = Nothing
	cn.Close
	Set cn = Nothing

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
