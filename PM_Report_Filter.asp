<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<script language="javascript">
</script>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>PM Report Filter</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="InspectionsFunctions.asp"-->
<link rel=STYLESHEET href='formstyle.css' type='text/css'>
<style>
</style>
<script language="Javascript">
function openhelp() {
 window.open("Equipment Inspection Database Users Guide.doc","userguide");
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
<!--#include file="datepicker.js"-->
</script>
</head>
<%
'*************
' Revision History
' 
' Keith Brooks - Monday, January 16, 2012
'   Creation.
'*************

Dim sqlString
Dim cn
Dim rs
Dim currentuser
Dim access
Dim start_date
Dim end_date
Dim equipment_item_id

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("inspections", "pm_report_filter", currentuser)
If access <> "none" Then

	Response.Write "<body>"
	Response.Write "<form id='form1' name='form1' action='pm_report.asp' method='post'>"
		
	'Define the ado connection and recordset objects.
	set cn = CreateObject("adodb.connection")
	cn.CursorLocation = 3
	cn.Open = DBString
	set rs = CreateObject("adodb.recordset")

	'Assign the session variable values to local variables, if they exist.
	start_date = Session("start_date")
	end_date = Session("end_date")
	equipment_item_id = Session("equipment_item_id")
	Session.Contents.Remove("start_date")
	Session.Contents.Remove("end_date")
	Session.Contents.Remove("equipment_item_id")
	
	'Draw the header.
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<td class='noprint' style='text-align:left;vertical-align:top;width:50%'><a href='default.asp'>Home</a></td>"
	Response.Write "<td class='noprint' style='text-align:right;vertical-align:top;width:50%'><a href='' onclick='openhelp();return false;' title='Open the User Guide'>Help</a></td>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	Response.Write "<td colspan='2' style='text-align:center;vertical-align:top;font-size:12pt;font-weight:bold'>Select Filter for PM Report</td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "<br />"
	
	Response.Write "<table style='width:100%;border:none'>"
	'Draw the start and end date fields and calendar buttons.
	Response.Write "<tr>"
	Response.Write "<td style='width:50%;text-align:right;padding-right:10px'>Start Date:&nbsp;&nbsp;"
	Response.Write "<input type='text' id='start_date' name='start_date' size='10' value='" & start_date & "' onchange='chkDate(this);' />"
	Response.Write "<a style='vertical-align:bottom' href='javascript: void(0);' onclick='displayDatePicker(""start_date"");return false;'><img src='../images/calendar.bmp' alt='Calendar' /></a></td>"
	Response.Write "<td style='width:50%;text-align:left;padding-left:10px'>End Date:&nbsp;&nbsp;"
	Response.Write "<input type='text' id='end_date' name='end_date' size='10' value='" & end_date & "' onchange='chkDate(this);' />"
	Response.Write "<a style='vertical-align:bottom' href='javascript: void(0);' onclick='displayDatePicker(""end_date"");return false;'><img src='../images/calendar.bmp' alt='Calendar' /></a></td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "<br />"
	Response.Write "<br />"
	
	'Draw the equipment item dropdown list.
	Response.Write "<div style='text-align:center'>Equipment Item:&nbsp;&nbsp;"
	Response.Write "<select style='background-color:white;font-size:10pt' id='equipment_item_id' name='equipment_item_id'>"
	Response.Write "<option value='' />"
	sqlString = "SELECT equipment_item_id,equipment_item_tag " & _
				"FROM equipment_items " & _
				"WHERE equipment_type_id=6"
	Set rs = cn.Execute(sqlString)
	If Not rs.BOF Then
		rs.MoveFirst
		Do While Not rs.EOF
			If IsNumeric(equipment_item_id) Then
				If CLng(rs(0)) = CLng(equipment_item_id) Then
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
	rs.Close
	Response.Write "</select>"
	Response.Write "</div>"
	Response.Write "<br />"
	Response.Write "<br />"
	Response.Write "<br />"
	
	'Draw the submit button.
	Response.Write "<div style='text-align:center'>"
	Response.Write "<input type='submit' id='submit1' name='submit1' value='Get Report' />"
	Response.Write "</div>"
	
	Set rs = Nothing
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
