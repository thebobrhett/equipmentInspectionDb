<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<script language="javascript">
function doSubmit() {
 document.getElementById('PleaseWait').style.display = 'block';
 document.form1.flowflag.value='false';
 document.form1.submit();
}
function doFind() {
 document.getElementById('PleaseWait').style.display = 'block';
 document.form1.submit();
}
function openhelp() {
 window.open("Equipment Inspections Database Administrators Guide.doc","userguide");
}
<!--#include file="datepicker.js"-->
</script>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Inspections Audit Trail</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="InspectionsFunctions.asp"-->
<link rel=STYLESHEET href='formstyle.css' type='text/css'>
</head>
<body>
<table style="width:100%;border:none">
	<tr>
		<td style="text-align:left;width:20%"><a href="adminmenu.asp">Menu</a></td>
		<td style="text-align:center;width:60%"><h1 />Inspections Audit Trail</td>
		<td style="text-align:right;width:20%"><a href="" onclick="openhelp();return false;" title="Open the Admin Guide">Help</a></td>
	</tr>
</table>
<form id="form1" name="form1" action="AuditTrail.asp" method="post">
<%
Dim sqlString
Dim cn
Dim rs
Dim criteria
Dim tagname
Dim tagdesc
Dim changeType
Dim changeTypes()

ReDim changeTypes(2)
changeTypes(0) = "delete"
changeTypes(1) = "insert"
changeTypes(2) = "update"


'Define the ado connection and recordset objects.
set cn = CreateObject("adodb.connection")
cn.Open = DBString
set rs = CreateObject("adodb.recordset")

'Draw "Please Wait..." message that will be displayed when this page is
'reloading, saving data, or moving to another page.
%>
	<div class="helptext" id="PleaseWait" style="display: none; text-align:center; color:White; vertical-align:top;border-style:none;position:absolute;top:0px;left:0px">
		<table id="MyTable" bgcolor="blue">
			<tr><td style="width: 95px"><b><font color="white">Please Wait...</font></b></td></tr>
		</table>
	</div>
<%
'Draw the criteria selection lists.
Response.Write "<table style='width:100%;border:none'>"
Response.Write "<tr>"
Response.Write "<th style='width:30%'>Date Range</th>"
Response.Write "<th style='width:30%'>Equipment Item</th>"
Response.Write "<th style='width:20%'>Type</th>"
Response.Write "<th style='width:30%'>Modifier</th>"
Response.Write "</tr>"
Response.Write "<tr>"

Response.Write "<td style='vertical-align:top'>"
Response.Write "<table style='width:100%;padding:5px'>"
'Draw the start date.
Response.Write "<tr>"
response.write "<td style='text-align:right;padding:10px'><b>Start Date: </b></td>"
If Request("start_date") <> "" Then
	Response.Write "<td style='text-align:left'><input type='text' class='text' name='start_date' size='10' value='" & Request("start_date") & "' onchange='checkDate_onchange(0)' />"
Else
	Response.Write "<td style='text-align:left'><input type='text' class='text' name='start_date' size='10' value='' onchange='checkDate_onchange(0)' />"
End If
Response.Write "<a href='javascript: displayDatePicker(""start_date"");'><img src='../images/calendar.bmp' alt='Calendar' style='vertical-align:top'></a></td>"
Response.Write "</tr>"
'Draw the end date.
Response.Write "<tr>"
response.write "<td style='text-align:right;padding:10px'><b>End Date: </b></td>"
If Request("end_date") <> "" Then
	Response.Write "<td style='text-align:left'><input type='text' class='text' name='end_date' size='10' value='" & Request("end_date") & "' onchange='checkDate_onchange(1)' />"
Else
	Response.Write "<td style='text-align:left'><input type='text' class='text' name='end_date' size='10' value='' onchange='checkDate_onchange(1)' />"
End If
Response.Write "<a href='javascript: displayDatePicker(""end_date"");'><img src='../images/calendar.bmp' alt='Calendar' style='vertical-align:top'></a></td>"
Response.Write "</tr>"
Response.Write "</table>"
Response.Write "</td>"

'Load the equipment types dropdown list.
Response.Write "<td style='vertical-align:top'>"
Response.Write "<table style='width:100%'>"
Response.Write "<tr>"
Response.Write "<tr>"
Response.Write "<td>Type: <select id='equiptype' name='equiptype' onchange='doSubmit();'>"
sqlString = "SELECT equipment_type_id,equipment_type_name FROM equipment_types ORDER BY equipment_type_name"
Set rs = cn.Execute(sqlString)
If Not rs.BOF Then
	rs.MoveFirst
	Response.Write "<option value=''> "
	Do While Not rs.EOF
		If Request("equiptype") <> "" Then
			If CLng(rs(0)) = CLng(Request("equiptype")) Then
				Response.Write "<option value='" & rs(0) & "' selected>" & rs(1)
			Else
				Response.Write "<option value='" & rs(0) & "'>" & rs(1)
			End If
		Else
			Response.Write "<option value='" & rs(0) & "'>" & rs(1)
		End If
		rs.MoveNext
	Loop
End If
rs.Close
Response.Write "</select></td>"
Response.Write "</tr>"
'Load the areas dropdown list.
Response.Write "<tr>"
Response.Write "<td>Area: <select id='area' name='area' onchange='doSubmit();'>"
sqlString = "SELECT DISTINCT area FROM equipment_items ORDER BY area"
Set rs = cn.Execute(sqlString)
If Not rs.BOF Then
	rs.MoveFirst
	Response.Write "<option value=''> "
	Do While Not rs.EOF
		If Request("area") <> "" Then
			If rs(0) = Request("area") Then
				Response.Write "<option value='" & rs(0) & "' selected>" & rs(0)
			Else
				Response.Write "<option value='" & rs(0) & "'>" & rs(0)
			End If
		Else
			Response.Write "<option value='" & rs(0) & "'>" & rs(0)
		End If
		rs.MoveNext
	Loop
End If
rs.Close
Response.Write "</select></td>"
Response.Write "</tr>"
'If equipment type has been selected, display a list of items.
If Request("equiptype") <> "" Then
	sqlString = "SELECT equipment_item_id,equipment_item_tag FROM equipment_items " & _
			"WHERE "
	If Request("equiptype") <> "" Then
		sqlString = sqlString & "equipment_type_id=" & Request("equiptype")
		If Request("area") <> "" Then
			sqlString = sqlString & " AND area='" & Request("area") & "'"
		End If
	Else
		sqlString = sqlString & "area='" & Request("area") & "'"
	End If
	sqlString = sqlString & " ORDER BY equipment_item_tag"
	Set rs = cn.Execute(sqlString)
	Response.Write "<tr>"
	Response.Write "<td><select id='item_id' name='item_id' size='10'>"
	If Not rs.BOF Then
		rs.MoveFirst
		Do While Not rs.EOF
			If Request("item_id") <> "" Then
				If CLng(Request("item_id")) = CLng(rs(0)) Then
					Response.Write "<option value='" & rs(0) & "' selected />" & rs(1)
				Else
					Response.Write "<option value='" & rs(0) & "' />" & rs(1)
				End If
			Else
				Response.Write "<option value='" & rs(0) & "' />" & rs(1)
			End If
			rs.MoveNext
		Loop
	End If
	Response.Write "</select></td>"
	Response.Write "</tr>"
End If
Response.Write "</table>"
Response.Write "</td>"

'Load the change type dropdown list.
Response.Write "<td style='vertical-align:top'><select name='changetype'>"
Response.Write "<option value=''> "
For Each changeType In changeTypes
	If Request("changetype") <> "" Then
		If changeType = Request("changetype") Then
			Response.Write "<option value='" & changeType & "' selected>" & changeType
		Else
			Response.Write "<option value='" & changeType & "'>" & changeType
		End If
	Else
		Response.Write "<option value='" & changeType & "'>" & changeType
	End If
Next
Response.Write "</select></td>"

'Load the modifier dropdown list.
Response.Write "<td style='vertical-align:top'><select name='modifier'>"
sqlString = "SELECT DISTINCT change_user FROM audit_trail ORDER BY change_user"
Set rs = cn.Execute(sqlString)
If Not rs.BOF Then
	rs.MoveFirst
	Response.Write "<option value=''> "
	Do While Not rs.EOF
		If Request("modifier") <> "" Then
			If rs(0) = Request("modifier") Then
				Response.Write "<option value='" & rs(0) & "' selected>" & rs(0)
			Else
				Response.Write "<option value='" & rs(0) & "'>" & rs(0)
			End If
		Else
			Response.Write "<option value='" & rs(0) & "'>" & rs(0)
		End If
		rs.MoveNext
	Loop
End If
rs.Close
Response.Write "</select></td>"
Response.Write "</tr>"
Response.Write "</table>"

Response.Write "<br />"
Response.Write "<table style='width:100%'>"
Response.Write "<tr>"
Response.Write "<td style='width:33%'>&nbsp;</td>"
Response.Write "<td style='width:34%;text-align:center'><input type='button' id='submit1' name='submit1' value='Find' style='font-size:10pt' onclick='doFind();'></td>"
Response.Write "<td style='width:33%'>&nbsp;</td>"
Response.Write "</tr>"
Response.Write "</table>"

'If any of the criteria have been selected, display the list box with the results.
criteria = ""
If Request("start_date") <> "" Then
	criteria = "change_date > '" & FormatMySQLDateTime(Request("start_date")) & "'"
End If
If Request("end_date") <> "" Then
	If criteria = "" Then
		criteria = "change_date < '" & FormatMySQLDateTime(DateAdd("d",1.0,Request("end_date"))) & "'"
	Else
		criteria = criteria & " AND change_date < '" & FormatMySQLDateTime(DateAdd("d",1.0,Request("end_date"))) & "'"
	End If
End If
If Request("equiptype") <> "" Then
	If criteria = "" Then
		criteria = "change_table='" & LCase(GetEquipmentTypeName(Request("equiptype"))) & "_inspection_data'"
	Else
		criteria = criteria & " AND change_table='" & Lcase(GetEquipmentTypeName(Request("equiptype"))) & "_inspection_data'"
	End If
End If
If Request("item_id") <> "" Then
	If criteria = "" Then
		criteria = "change_item_id=" & Request("item_id")
	Else
		criteria = criteria & " AND change_item_id=" & Request("item_id")
	End If
End If
If Request("changetype") <> "" Then
	If criteria = "" Then
		criteria = "change_type='" & Request("changetype") & "'"
	Else
		criteria = criteria & " AND change_type='" & Request("changetype") & "'"
	End If
End If
If Request("modifier") <> "" Then
	If criteria = "" Then
		criteria = "change_user='" & Request("modifier") & "'"
	Else
		criteria = criteria & " AND change_user='" & Request("modifier") & "'"
	End If
End If
If Request("flowflag") = "true" And criteria <> "" Then
	sqlString = "SELECT change_date,change_user,equipment_item_tag," & _
				"CONCAT(change_table,'.',change_field),old_value,new_value," & _
				"change_type " & _
				"FROM audit_trail LEFT JOIN equipment_items " & _
				"ON audit_trail.change_item_id=equipment_items.equipment_item_id " & _
				"WHERE " & criteria & " ORDER BY audit_trail_id"
	Dim returned
	Set rs = cn.Execute(sqlString,returned)
	Response.Write "<div style='text-align:center;color=blue'>"
	Response.Write returned & " records returned</div>"
	Response.Write "<table style='width:100%'>"
	Response.Write "<tr>"
	Response.Write "<th id='mediumth'>Timestamp</th>"
	Response.Write "<th id='mediumth'>Modifier</th>"
	Response.Write "<th id='mediumth'>Item</th>"
	Response.Write "<th id='mediumth'>Table.Field</th>"
	Response.Write "<th id='mediumth'>Old Value</th>"
	Response.Write "<th id='mediumth'>New Value</th>"
	Response.Write "<th id='mediumth'>Change Type</th>"
	Response.Write "</tr>"
	If Not rs.BOF Then
		rs.MoveFirst
		Do While Not rs.EOF
			Response.Write "<tr>"
			If Not IsNull(rs(0)) Then
				Response.Write "<td id='mediumtd'>" & rs(0) & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Not IsNull(rs(1)) Then
				Response.Write "<td id='mediumtd'>" & rs(1) & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Not IsNull(rs(2)) Then
				Response.Write "<td id='mediumtd'>" & rs(2) & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Not IsNull(rs(3)) Then
				Response.Write "<td id='mediumtd'>" & rs(3) & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Not IsNull(rs(4)) And rs(4) <> "" Then
				Response.Write "<td id='mediumtd'>" & rs(4) & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Not IsNull(rs(5)) And rs(5) <> " " And rs(5) <> "" Then
				Response.Write "<td id='mediumtd'>" & rs(5) & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Not IsNull(rs(6)) Then
				Response.Write "<td id='mediumtd'>" & rs(6) & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			Response.Write "</tr>"
			rs.MoveNext
		Loop
	End If
	rs.Close
	Response.Write "</table>"
End If

Response.Write "<input type='hidden' name='flowflag' id='flowflag' value='true' />"

Set rs = Nothing
cn.Close
Set cn = Nothing
%>
</form>
</body>
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