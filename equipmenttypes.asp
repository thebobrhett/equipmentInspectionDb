<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<script language="javascript">
function doDelete(id) {
	if (confirm("Are you sure you want to delete record number "+id+"?")) {
		document.form1.action="adminaction.asp?action=delete&RECORD="+id;
		document.form1.submit()
	}
}
function openhelp() {
 window.open("Equipment Inspections Database Administrators Guide.doc","userguide");
}
</script>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Administer Equipment Types</title>
<link rel=STYLESHEET href='formstyle.css' type='text/css'>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="InspectionsFunctions.asp"-->
</head>

<%
'*************
' Revision History
' 
' Keith Brooks - Saturday, February 19, 2011
'   Creation.
'*************

dim cn
dim rs
dim recordid
Dim currentuser
Dim access
Dim recid
Dim sqlString
Dim sortkey
Dim sortdir

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("inspections","equipmenttypes",currentuser)
If access <> "none" Then

	set cn = CreateObject("adodb.connection")
	cn.Open = DBString
	set rs = CreateObject("adodb.recordset")

	If session("err") <> "" And session("err") <> "NONE" Then
	  Response.Write "<body onload='document.form1." & session("err") & ".focus();'>"
	ElseIf session("focus") <> "" And session("focus") <> "NONE" Then
	  Response.Write "<body onload='document.form1." & session("focus") & ".focus();'>"
	  session("focus") = "NONE"
	Else
	  Response.Write "<body>"
	End If

	If request("record_id") <> "" Then
	  If IsNumeric(request("record_id")) Then
	    recordid = request("record_id")
	  Else
	    recordid = 0
	  End If
	Else
	  recordid = 0
	End If
	If request("sort") <> "" Then
		sortkey = request("sort")
	Else
		sortkey = "equipment_type_id"
	End If
	If request("direction") <> "" Then
		sortdir = request("direction")
	Else
		sortdir = "ASC"
	End If

	response.write "<table ID='headertable' width='100%'>"
	response.write "<tr>"
	response.write "<td ID='headertd' style='width:20%;text-align:left;vertical-align:top'><a href='adminmenu.asp' title='Open the administration main menu'>Menu</a></td>"
	response.write "<td ID='headertd' style='width:60%;text-align:center;vertical-align:center'><h1/>Edit Equipment Types</td>"
	response.write "<td ID='headertd' style='width:20%;text-align:right;vertical-align:top'><a href='#' onClick='openhelp();return false;' title='Open the Admin Guide'>Help</a></td>"
	response.write "</tr>"

	response.write "<tr>"
	response.write "<td id='headertd'>&nbsp;</td>"
	If access = "write" Or access = "delete" Then
		response.write "<td ID='headertd' style='text-align:center'><a href='equipmenttypes.asp?record_id=-1&sort=" & sortkey & "&direction=" & sortdir & "' title='Add a new equipment type record'>Add new record</a></td>"
	Else
		Response.Write "<td id='headertd'>&nbsp;</td>"
	End If
	response.write "<td id='headertd'>&nbsp;</td>"
	response.write "</tr>"
	response.write "</table>"

	Response.Write "<br />"

	'Draw the header
	Response.Write "<div style='text-align:center'>"
	response.Write "<table width='75%'>"
	response.Write "<tr>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='equipmenttypes.asp?sort=equipment_type_id&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Equipment Type ID&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='equipmenttypes.asp?sort=equipment_type_id&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='equipmenttypes.asp?sort=equipment_type_name&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Equipment Type Name&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='equipmenttypes.asp?sort=equipment_type_name&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='equipmenttypes.asp?sort=equipment_type_description&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Equipment Type Description&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='equipmenttypes.asp?sort=equipment_type_description&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='equipmenttypes.asp?sort=inspection_interval&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Inspection<br />Interval&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='equipmenttypes.asp?sort=inspection_interval&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='equipmenttypes.asp?sort=inspection_interval_units&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Inspection<br />Interval<br />Units&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='equipmenttypes.asp?sort=inspection_interval_units&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	If access = "write" Or access = "delete" Then
 	  response.Write "<th id='mediumth'>&nbsp;</th>"
 	End If
	If access = "delete" Or recordid < 0 Then
 	  response.Write "<th id='mediumth'>&nbsp;</th>"
 	End If
	response.Write "</tr>"

	response.Write "<form action='adminaction.asp' id='form1' method='post' name='form1'>"

	response.write "<input type='hidden' name='RECORD' value='" & recordid & "'>"
	response.write "<input type='hidden' name='SORT' value='" & sortkey & "'>"
	response.write "<input type='hidden' name='DIRECTION' value='" & sortdir & "'>"

	sqlString = "SELECT * FROM equipment_types " & _
				"ORDER BY " & sortkey & " " & sortdir
	set rs = cn.Execute(sqlString)

	If Not rs.BOF Then
	  rs.MoveFirst
	End If
	  
	'If recordid<0, the user has selected "Add new record" so insert a blank data entry line
	'at the top of the form.
	If access = "write" Or access = "delete" Then
		If recordid < 0 Then
			Response.Write "<tr>"

			response.write "<td id='mediumtd'>&nbsp;</td>"

			If session("err") = "equipment_type_name" Then
				If session("equipment_type_name") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='equipment_type_name' name='equipment_type_name' size='30' value='" & session("equipment_type_name") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='equipment_type_name' name='equipment_type_name' size='30' value=''></td>"
				End If
			Else
				If session("equipment_type_name") <> "" Then
					response.write "<td id='mediumtd'><input type='text' class='text' id='equipment_type_name' name='equipment_type_name' size='30' value='" & session("equipment_type_name") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' class='text' id='equipment_type_name' name='equipment_type_name' size='30' value=''></td>"
				End If
			End If

			If session("err") = "equipment_type_description" Then
				If session("equipment_type_description") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><textarea id='equipment_type_descripiton' name='equipment_type_description' cols='30' rows='2'>" & session("equipment_type_description") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><textarea id='equipment_type_description' name='equipment_type_description' cols='30' rows='2'></textarea></td>"
				End If
			Else
				If session("equipment_type_description") <> "" Then
					response.write "<td id='mediumtd'><textarea id='equipment_type_description' name='equipment_type_description' cols='30' rows='2'>" & session("equipment_type_description") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd'><textarea id='equipment_type_description' name='equipment_type_description' cols='30' rows='2'></textarea></td>"
				End If
			End If

			If session("err") = "inspection_interval" Then
				If session("inspection_interval") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='inspection_interval' name='inspection_interval' size='5' value='" & session("inspection_interval") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='inspection_interval' name='inspection_interval' size='5' value=''></td>"
				End If
			Else
				If session("inspection_interval") <> "" Then
					response.write "<td id='mediumtd'><input type='text' class='text' id='inspection_interval' name='inspection_interval' size='5' value='" & session("inspection_interval") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' class='text' id='inspection_interval' name='inspection_interval' size='5' value=''></td>"
				End If
			End If

			If session("err") = "inspection_interval_units" Then
				Response.Write "<td id='mediumtd' style='background-color:red'>"
			Else
				Response.Write "<td id='mediumtd'>"
			End If
			Response.Write "<select id='inspection_interval_units' name='inspection_interval_units'>"
			Response.Write "<option value=''>"
			If session("inspection_interval_units") = "days" Then
				Response.Write "<option value='days' selected>days"
			Else
				Response.Write "<option value='days'>days"
			End If
			If Session("inspection_interval_units") = "months" Then
				Response.Write "<option value='months' selected>months"
			Else
				Response.Write "<option value='months'>months"
			End If
			If Session("inspection_interval_units") = "years" Then
				Response.Write "<option value='years' selected>years"
			Else
				Response.Write "<option value='years'>years"
			End If
			Response.Write "</select></td>"

			response.write "<td id='mediumtd' style='text-align:center'><input type='submit' value='Submit' id='submit1' name='submit1' style='font-size:8pt;background-color:white' title='Apply changes to this record'></td>"

			Response.Write "<td id='mediumtd' style='text-align:center'><a href='equipmenttypes.asp?sort=" & sortkey & "&direction=" & sortdir & "' title='Cancel changes to this record'>Cancel</a></td>"

			Response.Write "</tr>"
		End If
	End If

	Do While Not rs.EOF
		Response.Write "<tr>"
		If CLng(rs("equipment_type_id")) = CLng(recordid) Then
			'Draw the data entry line
			response.write "<td id='mediumtd'>" & rs("equipment_type_id") & "</td>"

			If session("err") = "equipment_type_name" Then
				If session("equipment_type_name") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='equipment_type_name' name='equipment_type_name' size='30' value='" & session("equipment_type_name") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='equipment_type_name' name='equipment_type_name' size='30' value='" & rs("equipment_type_name") & "'></td>"
				End If
			Else
				If session("equipment_type_name") <> "" Then
					response.write "<td id='mediumtd'><input type='text' class='text' id='equipment_type_name' name='equipment_type_name' size='30' value='" & session("equipment_type_name") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' class='text' id='equipment_type_name' name='equipment_type_name' size='30' value='" & rs("equipment_type_name") & "'></td>"
				End If
			End If

			If session("err") = "equipment_type_description" Then
				If session("equipment_type_description") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><textarea id='equipment_type_description' name='equipment_type_description' cols='30' rows='2'>" & session("equipment_type_description") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><textarea id='equipment_type_description' name='equipment_type_description' cols='30' rows='2'>" & rs("equipment_type_description") & "</textarea></td>"
				End If
			Else
				If session("equipment_type_description") <> "" Then
					response.write "<td id='mediumtd'><textarea id='equipment_type_description' name='equipment_type_description' cols='30' rows='2'>" & session("equipment_type_description") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd'><textarea id='equipment_type_description' name='equipment_type_description' cols='30' rows='2'>" & rs("equipment_type_description") & "</textarea></td>"
				End If
			End If

			If session("err") = "inspection_interval" Then
				If session("inspection_interval") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='inspection_interval' name='inspection_interval' size='5' value='" & session("inspection_interval") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='inspection_interval' name='inspection_interval' size='5' value='" & rs("inspection_interval") & "'></td>"
				End If
			Else
				If session("inspection_interval") <> "" Then
					response.write "<td id='mediumtd'><input type='text' class='text' id='inspection_interval' name='inspection_interval' size='5' value='" & session("inspection_interval") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' class='text' id='inspection_interval' name='inspection_interval' size='5' value='" & rs("inspection_interval") & "'></td>"
				End If
			End If

			If session("err") = "inspection_interval_units" Then
				response.write "<td id='mediumtd' style='background-color:red'>"
			Else
				response.write "<td id='mediumtd'>"
			End If
			Response.Write "<select id='inspection_interval_units' name='inspection_interval_units'>"
			Response.Write "<option value=''>"
			If rs("inspection_interval_units") = "days" Or session("inspection_interval_units") = "days" Then
				Response.Write "<option value='days' selected>days"
			Else
				Response.Write "<option value='days'>days"
			End If
			If rs("inspection_interval_units") = "months" Or Session("inspection_interval_units") = "months" Then
				Response.Write "<option value='months' selected>months"
			Else
				Response.Write "<option value='months'>months"
			End If
			If rs("inspection_interval_units") = "years" Or Session("inspection_interval_units") = "years" Then
				Response.Write "<option value='years' selected>years"
			Else
				Response.Write "<option value='years'>years"
			End If
			Response.Write "</select></td>"

			If access = "write" Or access = "delete" Then
				response.write "<td id='mediumtd' style='text-align:center'><input type='submit' value='Submit' id='submit1' name='submit1' style='font-size:8pt;background-color:white' title='Apply changes to this record'></td>"
			End If

			Response.Write "<td id='mediumtd' style='text-align:center'><a href='equipmenttypes.asp?sort=" & sortkey & "&direction=" & sortdir & "' title='Cancel changes to this record'>Cancel</a></td>"
'			If access = "delete" Then
'				recid = rs("equipment_type_id")
'				response.write "<td id='mediumtd'><a href='javascript:doDelete(" & recid & ");' title='Delete this record'>Delete</a></td>"
'			End If
		Else
			'Draw the history records
			response.write "<tr>"
			response.write "<td id='mediumtd'>" & rs("equipment_type_id") & "</td>"
			If rs("equipment_type_name") <> "" Then
				response.write "<td id='mediumtd'>" & rs("equipment_type_name") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If rs("equipment_type_description") <> "" Then
				response.write "<td id='mediumtd'>" & rs("equipment_type_description") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If rs("inspection_interval") <> "" Then
				response.write "<td id='mediumtd'>" & rs("inspection_interval") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If rs("inspection_interval_units") <> "" Then
				response.write "<td id='mediumtd'>" & rs("inspection_interval_units") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If access = "write" Or access = "delete" Then
				response.write "<td id='mediumtd'><a href='equipmenttypes.asp?record_id=" & rs("equipment_type_id") & "&sort=" & sortkey & "&direction=" & sortdir & "' title='Edit this record'>Edit</a></td>"
			End If
			If access = "delete" Then
				recid = rs("equipment_type_id")
				response.write "<td id='mediumtd'><a href='javascript:doDelete(" & recid & ");' title='Delete this record'>Delete</a></td>"
			ElseIf recordid < 0 Then
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			response.write "</tr>"
		End If
		rs.Movenext
	loop
	rs.Close
	
	Response.Write "</form>"
	Response.Write "</table>"
	Response.Write "</div>"
	Response.Write "</body>"

	'session("err") = "NONE"

	Set rs = Nothing
	cn.Close
	Set cn = Nothing

Else
	response.write "<h1>You don't have permission to access this page.</h1>"
	response.write "<br />"
	response.write "<a href='" & request("http_referer") & "'>Return to previous page</a>"
End If
%>
</html>
