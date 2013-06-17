<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<html>
<head>
<script language="javascript">
var needToConfirm = false;

function openhelp() {
 window.open("Equipment Inspection Database Administrators Guide.doc","userguide");
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
 window.location.href='tank_technicaldata.asp?itemid='+document.form1.itemID.value+'&edit=false';
}
function editmode() {
 //Put the page in edit mode.
 window.location.href='tank_technicaldata.asp?itemid='+document.form1.itemID.value+'&edit=true';
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
<title>Tank Technical Data</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="InspectionsFunctions.asp"-->
<link rel=STYLESHEET href='formstyle.css' type='text/css'>
</head>
<%
'*************
' Revision History
' 
' Keith Brooks - Monday, March 7, 2011
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
Dim inspection_standard
Dim next_inspection_date
Dim inspection_frequency
Dim inspection_frequency_units
Dim previous_inspection_date

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("inspections", "tank_technicaldata", currentuser)
If access <> "none" Then

	If LCase(Request("edit")) = "true" And (access = "write" Or access = "delete") Then
		editMode = True
		field_disabled = ""
		Response.Write "<body  onload='document.form1.state_number.focus();'>"
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
	Response.Write "<input type='hidden' id='equipType' name='equipType' value='tank' />"
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
			Response.Write "<td style='text-align:left;vertical-align:top;font-size:12pt;font-weight:bold'>Vessel - Technical Data Sheet</td>"
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
			sqlString = "SELECT * FROM tank_technical_data " & _
					"WHERE equipment_item_id=" & Request("itemID")
			Set rs = cn.Execute(sqlString)
			If Not rs.BOF Then
				'If the record exists, assign the existing values to the variables.
				rs.MoveFirst
				technical_data_id = rs("technical_data_id")
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
				inspection_standard = rs("inspection_standard")
				next_inspection_date = rs("next_inspection_date")
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
				
			Else
				'If the record doesn't exist, initialize the variables.
				technical_data_id = 0
				state_number = ""
				relief_device = ""
				relief_device_pressure = ""
				relief_device_pressure_units = ""
				lethal_service = ""
				capacity = ""
				capacity_units = ""
				weight_empty = ""
				weight_empty_units = ""
				height_length = ""
				height_length_units = ""
				inside_diameter = ""
				inside_diameter_units = ""
				shell_material = ""
				shell_thickness = ""
				shell_thickness_units = ""
				shell_min_thickness = ""
				shell_min_thickness_units = ""
				head_material = ""
				head_thickness = ""
				head_thickness_units = ""
				head_min_thickness = ""
				head_min_thickness_units = ""
				lining_material = ""
				lining_thickness = ""
				lining_thickness_units = ""
				jacket_material = ""
				jacket_thickness = ""
				jacket_thickness_units = ""
				mawp = ""
				mawp_units = ""
				shell_test_press = ""
				shell_test_press_units = ""
				jacket_test_press = ""
				jacket_test_press_units = ""
				date_built = ""
				national_board_number = ""
				lining_mfgr = ""
				manufacturer = ""
				mfgr_serial_number = ""
				drawing_number = ""
				jacket_type_description = ""
				inspection_standard = ""
				inspection_frequency = 0
				inspection_frequency_units = ""
				next_inspection_date = ""
			
			End If
			
			'If inspection frequency or its units are not specified, look the default values
			'up in the equipment types table.
			If inspection_frequency = 0 Or inspection_frequency_units = "" Then
				sqlString = "SELECT inspection_interval,inspection_interval_units " & _
						"FROM equipment_types WHERE equipment_type_name='Tank'"
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
					"FROM tank_inspection_data " & _
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
			Response.Write "<div style='text-align:left'>"
			Response.Write "<table style='width:50%;border:none'>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd' style='width:45%'>Relief Device:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='relief_device' name='relief_device' value='" & relief_device & "' onclick='setupdate();' /></td>"
			Else
				Response.Write relief_device & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Relief Device Pressure:</td>"
			Response.Write "<td id='techtd' style='width:20%'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='relief_device_pressure' name='relief_device_pressure' value='" & relief_device_pressure & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write relief_device_pressure & "</td>"
			End If
			Response.Write "<td id='techtd' style='width:35%'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='relief_device_pressure_units' name='relief_device_pressure_units' value='" & relief_device_pressure_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write relief_device_pressure_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Lethal Service:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='lethal_service' name='lethal_service' value='" & lethal_service & "' onchange='setupdate();' /></td>"
			Else
				Response.Write lethal_service & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Capacity:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='capacity' name='capacity' value='" & capacity & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write capacity & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='capacity_units' name='capacity_units' value='" & capacity_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write capacity_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Weight Empty:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='weight_empty' name='weight_empty' value='" & weight_empty & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write weight_empty & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='weight_empty_units' name='weight_empty_units' value='" & weight_empty_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write weight_empty_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Height/Length:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='height_length' name='height_length' value='" & height_length & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write height_length & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='height_length_units' name='height_length_units' value='" & height_length_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write height_length_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Inside Diameter:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='inside_diameter' name='inside_diameter' value='" & inside_diameter & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write inside_diameter & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='inside_diameter_units' name='inside_diameter_units' value='" & inside_diameter_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write inside_diameter_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Shell Material:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='shell_material' name='shell_material' value='" & shell_material & "' onchange='setupdate();' /></td>"
			Else
				Response.Write shell_material & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Shell Thickness:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='shell_thickness' name='shell_thickness' value='" & shell_thickness & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write shell_thickness & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='shell_thickness_units' name='shell_thickness_units' value='" & shell_thickness_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write shell_thickness_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Shell Min. Thickness:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='shell_min_thickness' name='shell_min_thickness' value='" & shell_min_thickness & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write shell_min_thickness & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='shell_min_thickness_units' name='shell_min_thickness_units' value='" & shell_min_thickness_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write shell_min_thickness_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Head Material:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='head_material' name='head_material' value='" & head_material & "' onchange='setupdate();' /></td>"
			Else
				Response.Write head_material & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Head Thickness:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='head_thickness' name='head_thickness' value='" & head_thickness & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write head_thickness & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='head_thickness_units' name='head_thickness_units' value='" & head_thickness_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write head_thickness_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Head Min. Thickness:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='head_min_thickness' name='head_min_thickness' value='" & head_min_thickness & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write head_min_thickness & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='head_min_thickness_units' name='head_min_thickness_units' value='" & head_min_thickness_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write head_min_thickness_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Lining Material:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='lining_material' name='lining_material' value='" & lining_material & "' onchange='setupdate();' /></td>"
			Else
				Response.Write lining_material & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Lining Thickness:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='lining_thickness' name='lining_thickness' value='" & lining_thickness & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write lining_thickness & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='lining_thickness_units' name='lining_thickness_units' value='" & lining_thickness_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write lining_thickness_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Jacket Material:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='jacket_material' name='jacket_material' value='" & jacket_material & "' onchange='setupdate();' /></td>"
			Else
				Response.Write jacket_material & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Jacket Thickness:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='jacket_thickness' name='jacket_thickness' value='" & jacket_thickness & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write jacket_thickness & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='jacket_thickness_units' name='jacket_thickness_units' value='" & jacket_thickness_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write jacket_thickness_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>MAWP:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='mawp' name='mawp' value='" & mawp & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write mawp & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='mawp_units' name='mawp_units' value='" & mawp_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write mawp_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Shell Test Press:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='shell_test_press' name='shell_test_press' value='" & shell_test_press & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write shell_test_press & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%'  id='shell_test_press_units' name='shell_test_press_units' value='" & shell_test_press_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write shell_test_press_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Jacket Test Press:</td>"
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='jacket_test_press' name='jacket_test_press' value='" & jacket_test_press & "' onchange='chkNumeric(this);setupdate();' /></td>"
			Else
				Response.Write jacket_test_press & "</td>"
			End If
			Response.Write "<td id='techtd'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='jacket_test_press_units' name='jacket_test_press_units' value='" & jacket_test_press_units & "' onchange='setupdate();' /></td>"
			Else
				Response.Write jacket_test_press_units & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Year Built:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='date_built' name='date_built' value='" & date_built & "' onchange='setupdate();' />"
			Else
				Response.Write date_built & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>National Board No:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='national_board_number' name='national_board_number' value='" & national_board_number & "' onchange='setupdate();' /></td>"
			Else
				Response.Write national_board_number & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Lining MFGR:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='lining_mfgr' name='lining_mfgr' value='" & lining_mfgr & "' onchange='setupdate();' /></td>"
			Else
				Response.Write lining_mfgr & "</td>"
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
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>MFGR Serial No:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='mfgr_serial_number' name='mfgr_serial_number' value='" & mfgr_serial_number & "' onchange='setupdate();' /></td>"
			Else
				Response.Write mfgr_serial_number & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Drawing No:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='drawing_number' name='drawing_number' value='" & drawing_number & "' onchange='setupdate();' /></td>"
			Else
				Response.Write drawing_number & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Jacket Type Desc:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='jacket_type_description' name='jacket_type_description' value='" & jacket_type_description & "' onchange='setupdate();' /></td>"
			Else
				Response.Write jacket_type_description & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Inspection Standard:</td>"
			Response.Write "<td id='techtd' colspan='2'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='inspection_standard' name='inspection_standard' value='" & inspection_standard & "' onchange='setupdate();' /></td>"
			Else
				Response.Write inspection_standard & "</td>"
			End If
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd'>Inspection Freq:</td>"
			Response.Write "<td id='techtd' style='width:20%'>"
			If editMode = True Then
				Response.Write "<input type='text' class='text' style='width:100%' id='inspection_frequency' name='inspection_frequency' value='" & inspection_frequency & "' onchange='chkNumeric(this);addDate(""" & previous_inspection_date & """,document.form1.inspection_frequency.value,document.form1.inspection_frequency_units.value);return false;' /></td>"
			Else
				Response.Write inspection_frequency & "</td>"
			End If
			Response.Write "<td id='techtd' style='width:40%'>"
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
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td id='techtd colspan='6'>&nbsp;</td>"
			Response.Write "</tr>"
			Response.Write "</table>"
			Response.Write "</div>"
			If editMode = True Then
				Response.Write "<table style='width:100%;border:none'>"
				Response.Write "<tr>"
				Response.Write "<td id='techtd' style='width:50%;text-align:right'>"
				Response.Write "<button type='button' class='noprint' id='cancel' name='cancel' onclick='canceledit();return false;'>View</button></td>"
				Response.Write "<td id='techtd' style='width:50%;text-align:left'>"
				Response.Write "<button type='button' class='noprint' id='submit1' name='submit1' onclick='saveData();'>Submit</button></td>"
				Response.Write "</tr>"
				Response.Write "</table>"
			ElseIf access = "write" Or access = "delete" Then
				Response.Write "<table style='width:100%;border:none'>"
				Response.Write "<tr>"
				Response.Write "<td id='techtd' style='width:100%;text-align:center'>"
				Response.Write "<input type='button' class='noprint' id='edit' name='edit' value='Edit' onclick='editmode();return false;' /></td>"
				Response.Write "</tr>"
				Response.Write "</table>"
			End If
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
