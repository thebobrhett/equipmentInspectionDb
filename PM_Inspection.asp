<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
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
td {border:1px solid black;
	text-align:left;
	padding:3px;
	font-size:7pt}
th {vertical-align:bottom;
	font-size:7pt}
input {font-size:7pt}
button {font-size:7pt}
</style>
<script language="Javascript">
var needToConfirm = false;

function saveData(frm) {
 needToConfirm=false;
 frm.submit();
}
function setupdate() {
 needToConfirm=true;
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
function warn() {
 if (needToConfirm==true) {
  return "You have changed the data on this form and not submitted it.";
 }
}
window.onbeforeunload = warn;
<!--#include file="datepicker.js"-->
</script>
</head>
<%
'*************
' Revision History
' 
' Keith Brooks - Wednesday, January 4, 2012
'   Creation.
'*************

Dim sqlString
Dim cn
Dim rs
Dim rs2
Dim rs3
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
Dim pm_date
Dim in_need_of_repair
Dim seal_condition
Dim comments
Dim equipment_item_id
Dim pm_data_id
Dim equipment_items_subitem1_id
Dim equipment_items_subitem2_id
Dim pm_subitem1_data_id
Dim pm_subitem2_data_id
Dim front_bearing_db_level
Dim rear_bearing_db_level
Dim top_db_level
Dim spare1
Dim spare2
Dim startDate
Dim endDate
Dim tabIdx
Dim errVar
Dim errRecord
Dim pm_dateField

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("inspections", "pm_inspection", currentuser)
If access <> "none" Then

	errVar = Session("errVar")
	If IsNumeric(Session("errRecord")) Then
		errRecord = Session("errRecord")
	Else
		errRecord = 0
	End If
	Session.Contents.Remove("errVar")
	Session.Contents.Remove("errRecord")
	
	Response.Write "<body>"
		
	'Define the ado connection and recordset objects.
	set cn = CreateObject("adodb.connection")
	cn.CursorLocation = 3
	cn.Open = DBString
	set rs = CreateObject("adodb.recordset")
	Set rs2 = CreateObject("adodb.recordset")
	Set rs3 = CreateObject("adodb.recordset")

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
	Response.Write "<td colspan='5' class='noprint' style='border:none;text-align:left;vertical-align:top;width:50%'><a href='default.asp'>Home</a></td>"
	Response.Write "<td colspan='6' class='noprint' style='border:none;text-align:right;vertical-align:top;width:50%'><a href='' onclick='openhelp();return false;' title='Open the User Guide'>Help</a></td>"
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
	
	'Determine the start and end dates for the current quarter.
	If Month(Date) < 4 Then
		startDate = "1/1/" & Year(Date)
		endDate = "3/31/" & Year(Date)
	ElseIf Month(Date) < 7 Then
		startDate = "4/1/" & Year(Date)
		endDate = "6/30/" & Year(Date)
	ElseIf Month(Date) < 10 Then
		startDate = "7/1/" & Year(Date)
		endDate = "9/30/" & Year(Date)
	Else
		startDate = "10/1/" & Year(Date)
		endDate = "12/31/" & Year(Date)
	End If
	
	'Get the information for the equipment items and subitems.
	sqlString = "SELECT equipment_item_tag,equipment_item_description," & _
			"pm_priority,equipment_item_id " & _
			"FROM equipment_items " & _
			"WHERE equipment_type_id=6 " & _
			"ORDER BY equipment_item_id"
	Set rs = cn.Execute(sqlString)
	tabIdx = 0
	If Not rs.BOF Then
		rs.MoveFirst
		Do While Not rs.EOF
			tabIdx = tabIdx + 100
			'Fill in the variables.
			equipment_item_tag = rs("equipment_item_tag")
			equipment_item_description = rs("equipment_item_description")
			pm_priority = rs("pm_priority")
			equipment_item_id = rs("equipment_item_id")
			'Create the pm_date field name to allow the datepicker to work.
			pm_dateField = "pm_date" & equipment_item_id
			'Start the form tag for this record.
			Response.Write "<form id='form" & equipment_item_id & "' name='form" & equipment_item_id & "' action='PMAction.asp' method='post'>"
			
			'Get the subitems for this item.
			sqlString = "SELECT equipment_items_subitem_description, " & _
			"test_front_bearing,test_rear_bearing,test_top, " & _
			"equipment_items_subitem_id " & _
			"FROM equipment_items_subitems " & _
			"WHERE equipment_item_id=" & equipment_item_id & _
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
					If count = 1 Then
						equipment_items_subitem1_id = rs2("equipment_items_subitem_id")
					Else
						equipment_items_subitem2_id = rs2("equipment_items_subitem_id")
					End If
					
					'Fill in the pm data variables.  If the equipment_item_id is
					'the same as the errRecord, assign the session variables to
					'maintain entered data; otherwise, try to get the most recent
					'database values in the current quarter.
					If CLng(errRecord) = CLng(equipment_item_id) Then
						pm_date = Session(pm_dateField)
						in_need_of_repair = CBool(Session("in_need_of_repair"))
						seal_condition = Session("seal_condition")
						comments = Session("comments")
						front_bearing_db_level = Session("front_bearing_db_level")
						rear_bearing_db_level = Session("rear_bearing_db_level")
						top_db_level = Session("top_db_level")
						spare1 = CBool(Session("spare1"))
						spare2 = CBool(Session("spare2"))
						pm_data_id = Session("pm_data_id")
						pm_subitem1_data_id = Session("pm_subitem1_data_id")
						pm_subitem2_data_id = Session("pm_subitem2_data_id")
					Else
						'Get the pm data for this quarter.  If a record exists,
						'fill the data into the variables; otherwise, initialize them.
						sqlString = "SELECT pm_date,in_need_of_repair,seal_condition," & _
							"comments,front_bearing_db_level,rear_bearing_db_level," & _
							"top_db_level,spare,p.pm_data_id,pm_subitem_data_id " & _
							"FROM pm_data p LEFT JOIN pm_subitem_data s " & _
							"ON p.pm_data_id=s.pm_data_id " & _
							"WHERE equipment_items_subitem_id=" & rs2("equipment_items_subitem_id") & _
							" AND pm_date>='" & FormatMySQLDate(startDate) & "' " & _
							"AND pm_date<='" & FormatMySQLDate(endDate) & "' " & _
							"ORDER BY pm_date"
						Set rs3 = cn.Execute(sqlString)
						If Not rs3.BOF Then
							rs3.MoveFirst
							'Loop through to get the data for the most recent date.
							Do While Not rs3.EOF
								pm_date = rs3("pm_date")
								If IsNull(rs3("in_need_of_repair")) Then
									in_need_of_repair = 0
								Else
									in_need_of_repair = CBool(rs3("in_need_of_repair"))
								End If
								seal_condition = rs3("seal_condition")
								comments = rs3("comments")
								front_bearing_db_level = rs3("front_bearing_db_level")
								rear_bearing_db_level = rs3("rear_bearing_db_level")
								top_db_level = rs3("top_db_level")
								If count = 1 Then
									If IsNull(rs3("spare")) Then
										spare1 = 0
									Else
										spare1 = CBool(rs3("spare"))
									End If
								Else
									If IsNull(rs3("spare")) Then
										spare2 = 0
									Else
										spare2 = CBool(rs3("spare"))
									End If
								End If
								pm_data_id = rs3("pm_data_id")
								If count = 1 Then
									pm_subitem1_data_id = rs3("pm_subitem_data_id")
								Else
									pm_subitem2_data_id = rs3("pm_subitem_data_id")
								End If
								rs3.MoveNext
							Loop
						Else
							pm_date = ""
							in_need_of_repair = False
							seal_condition = ""
							comments = ""
							front_bearing_db_level = ""
							rear_bearing_db_level = ""
							top_db_level = ""
							spare1 = False
							spare2 = False
							pm_data_id = -1
							pm_subitem1_data_id = -1
							pm_subitem2_data_id = -1
						End If
						rs3.Close
					End If
					Response.Write "<tr>"
					If count = 1 Then
						Response.Write "<td rowspan='" & rowCount + 1 & "' style='text-align:center'>" & equipment_item_tag & "</td>"
						Response.Write "<input type='hidden' id='pm_data_id' name='pm_data_id' value='" & pm_data_id & "' />"
						Response.Write "<input type='hidden' id='equipment_item_id' name='equipment_item_id' value='" & equipment_item_id & "' />"
					End If
					If count = 1 And test_front_bearing And test_rear_bearing Then
						Response.Write "<td rowspan='2'>" & equipment_items_subitem_description & "</td>"
						Response.Write "<input type='hidden' id='pm_subitem1_data_id' name='pm_subitem1_data_id' value='" & pm_subitem1_data_id & "' />"
						Response.Write "<input type='hidden' id='equipment_items_subitem1_id' name='equipment_items_subitem1_id' value='" & equipment_items_subitem1_id & "' />"
					ElseIf count = 2 And test_top Then
						Response.Write "<td>" & equipment_items_subitem_description & "</td>"
						Response.Write "<input type='hidden' id='pm_subitem2_data_id' name='pm_subitem2_data_id' value='" & pm_subitem2_data_id & "' />"
						Response.Write "<input type='hidden' id='equipment_items_subitem2_id' name='equipment_items_subitem2_id' value='" & equipment_items_subitem2_id & "' />"
					End If
					If count = 1 And test_front_bearing Then
						Response.Write "<td>front bearing</td>"
					ElseIf count = 2 And test_top Then
						Response.Write "<td>top</td>"
					End If
					If count = 1 Then
						Response.Write "<td rowspan='" & rowCount + 1 & "'  style='text-align:center'>"
						If CLng(errRecord) = CLng(equipment_item_id) And errVar = pm_dateField Then
							Response.Write "<input type='text' style='background-color:lightpink' id='" & pm_dateField & "' name='" & pm_dateField & "' size='10' value='" & pm_date & "' tabindex='" & tabIdx+1 & "' onchange='chkDate(this);setupdate();' />"
						Else
							Response.Write "<input type='text' id='" & pm_dateField & "' name='" & pm_dateField & "' size='10' value='" & pm_date & "' tabindex='" & tabIdx+1 & "' onchange='chkDate(this);setupdate();' />"
						End If
						Response.Write "<br /><a href='javascript: void(0);' onclick='displayDatePicker(""" & pm_dateField & """);setupdate();return false;'><img src='../images/calendar.bmp' alt='Calendar'></a></td>"
						Response.Write "</td>"
					End If
					If count = 1 And test_front_bearing Then
						Response.Write "<td>"
						If CLng(errRecord) = CLng(equipment_item_id) And errVar = "front_bearing_db_level" Then
							Response.Write "<input type='text' style='text-align:right;background-color:lightpink' id='front_bearing_db_level' name='front_bearing_db_level' size='4' value='" & front_bearing_db_level & "' tabindex='" & tabIdx+2 & "' onchange='setupdate();' />"
						Else
							Response.Write "<input type='text' style='text-align:right' id='front_bearing_db_level' name='front_bearing_db_level' size='4' value='" & front_bearing_db_level & "' tabindex='" & tabIdx+2 & "' onchange='setupdate();' />"
						End If
						Response.Write "</td>"
					ElseIf count = 2 And test_top Then
						Response.Write "<td>"
						If CLng(errRecord) = CLng(equipment_item_id) And errVar = "top_db_level" Then
							Response.Write "<input type='text' style='text-align:right;background-color:lightpink' id='top_db_level' name='top_db_level' size='4' value='" & top_db_level & "' tabindex='" & tabIdx+4 & "' onchange='setupdate();' />"
						Else
							Response.Write "<input type='text' style='text-align:right' id='top_db_level' name='top_db_level' size='4' value='" & top_db_level & "' tabindex='" & tabIdx+4 & "' onchange='setupdate();' />"
						End If
						Response.Write "</td>"
					End If
					If count = 1 Then
						Response.Write "<td rowspan='" & rowCount + 1 & "' style='text-align:center'>"
						If in_need_of_repair Then
							Response.Write "<input type='checkbox' id='in_need_of_repair' name='in_need_of_repair' value='1' checked tabindex='" & tabIdx+5 & "' onchange='setupdate();' />"
						Else
							Response.Write "<input type='checkbox' id='in_need_of_repair' name='in_need_of_repair' value='1' tabindex='" & tabIdx+5 & "' onchange='setupdate();' />"
						End If
						Response.Write "</td>"
						Response.Write "<td rowspan='" & rowCount + 1 & "'>"
						Response.Write "<select id='seal_condition' name='seal_condition' tabindex='" & tabIdx+6 & "' onchange='setupdate();'>"
						Response.Write "<option value='' />"
						If seal_condition = "good" Then
							Response.Write "<option value='good' selected />good"
						Else
							Response.Write "<option value='good' />good"
						End If
						If seal_condition = "leaking" Then
							Response.Write "<option value='leaking' selected />leaking"
						Else
							Response.Write "<option value='leaking' />leaking"
						End If
						Response.Write "</select>"
						Response.Write "</td>"
						Response.Write "<td rowspan='" & rowCount + 1 & "'>"
						Response.Write "<textarea id='comments' name='comments' rows='4' cols='25' tabindex='" & tabIdx+7 & "' onchange='setupdate();'>" & comments & "</textarea>"
						Response.Write "</td>"
						Response.Write "<td rowspan='" & rowCount + 1 & "' style='text-align:center'>" & pm_priority & "</td>"
					End If
					If count = 1 And test_front_bearing And test_rear_bearing Then
						Response.Write "<td rowspan='2'>"
						If spare1 Then
							Response.Write "<input type='checkbox' id='spare1' name='spare1' value='1' checked tabindex='" & tabIdx+8 & "' onchange='setupdate();' />"
						Else
							Response.Write "<input type='checkbox' id='spare1' name='spare1' value='1' tabindex='" & tabIdx+8 & "' onchange='setupdate();' />"
						End If
						Response.Write "</td>"
					ElseIf count = 2 And test_top Then
						Response.Write "<td>"
						If spare2 Then
							Response.Write "<input type='checkbox' id='spare2' name='spare2' value='1' checked tabindex='" & tabIdx+9 & "' onchange='setupdate();' />"
						Else
							Response.Write "<input type='checkbox' id='spare2' name='spare2' value='1' tabindex='" & tabIdx+9 & "' onchange='setupdate();' />"
						End If
						Response.Write "</td>"
					End If
					If count = 1 Then
						Response.Write "<td rowspan='" & rowCount & "' style='border:none'>" & equipment_item_description & "</td>"
					Else
						Response.Write "<td style='text-align:left;border:none'>"
						If access = "write" Or access = "delete" Then
							Response.Write "<button type='button' id='submit1' name='submit1' tabindex='" & tabIdx+10 & "' onclick='saveData(this.form);'>Submit</button>"
						Else
							Response.Write "&nbsp;"
						End If
						Response.Write "</td>"
					End If
					Response.Write "</tr>"
					If count = 1 And test_rear_bearing Then
						Response.Write "<tr>"
						Response.Write "<td>rear bearing</td>"
						Response.Write "<td>"
						If CLng(errRecord) = CLng(equipment_item_id) And errVar = "rear_bearing_db_level" Then
							Response.Write "<input type='text' style='text-align:right;background-color:lightpink' id='rear_bearing_db_level' name='rear_bearing_db_level' size='4' value='" & rear_bearing_db_level & "' tabindex='" & tabIdx+3 & "' onchange='setupdate();' />"
						Else
							Response.Write "<input type='text' style='text-align:right' id='rear_bearing_db_level' name='rear_bearing_db_level' size='4' value='" & rear_bearing_db_level & "' tabindex='" & tabIdx+3 & "' onchange='setupdate();' />"
						End If
						Response.Write "</td>"
						If rowCount = 1 Then
							Response.Write "<td style='text-align:left;border:none'>"
							If access = "write" Or access = "delete" Then
								Response.Write "<button type='button' id='submit1' name='submit1' tabindex='" & tabIdx+10 & "' onclick='saveData(this.form);'>Submit</button>"
							Else
								Response.Write "&nbsp;"
							End If
							Response.Write "</td>"
						End If
						Response.Write "</tr>"
					End If
					
					rs2.MoveNext
				Loop
				Response.Write "</form>"
			End If
			rs2.Close
			rs.MoveNext
		Loop
	End If
	rs.Close
	
	Set rs = Nothing
	Set rs2 = Nothing
	Set rs3 = Nothing
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
