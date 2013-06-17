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
<title>Administer User Access</title>
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
Dim cn2
dim rs
Dim rs2
Dim rs3
dim recordid
dim name
Dim form
Dim access
Dim recid
Dim sortkey
Dim sortdir

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("inspections","useraccess",currentuser)
If access <> "none" Then

	set cn = CreateObject("adodb.connection")
	cn.Open = "driver={MySQL ODBC 3.51 Driver};option=16387;server=richmond.aksa.local;User=assetmgtuser;password=asset;DATABASE=asset_management;"
	Set rs = CreateObject("adodb.recordset")
	Set rs3 = CreateObject("adodb.recordset")
	Set cn2 = CreateObject("adodb.connection")
	cn2.Open = "driver={MySQL ODBC 3.51 Driver};option=16387;server=richmond.aksa.local;User=cipt_user;password=easy;DATABASE=cipt;"
	set rs2 = CreateObject("adodb.recordset")

	'If specified, clear the session variables.
	If Request.QueryString("clear") <> "" Then
		Session.Contents.RemoveAll
	End If
	
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
		sortkey = "form_name"
	End If
	If request("direction") <> "" Then
		sortdir = request("direction")
	Else
		sortdir = "ASC"
	End If

	response.write "<table ID='headertable' width='100%'>"
	response.write "<tr>"
	response.write "<td ID='headertd' style='width:30%;text-align:left;vertical-align:top'><a href='adminmenu.asp' title='Open the administration main menu'>Menu</a></td>"
	response.write "<td ID='headertd' style='width:40%;text-align:center'><h1/>Maintain User Access</td>"
	response.write "<td ID='headertd' style='width:30%;text-align:right;vertical-align:top'><a href='#' onClick='openhelp();return false;' title='Open the Admin Guide'>Help</a></td>"
	response.write "</tr>"

	response.write "<tr>"
	response.write "<td id='headertd'>&nbsp;</td>"
	If access = "write" Or access = "delete" Then
		response.write "<td ID='headertd' style='text-align:center'><a href='useraccess.asp?record_id=-1&sort=" & sortkey & "&direction=" & sortdir & "' title='Add a new user access record'>Add new record</a></td>"
	Else
		Response.Write "<td id='headertd'>&nbsp;</td>"
	End If
	response.write "<td id='headertd'>&nbsp;</td>"
	response.write "</tr>"
	response.write "</table>"

	'Draw the header
	response.Write "<table align='center' width='90%'>"
	response.Write "<tr>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='useraccess.asp?sort=form_name&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Form Name&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='useraccess.asp?sort=form_name&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='useraccess.asp?sort=role_name&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Role&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='useraccess.asp?sort=role_name&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='useraccess.asp?sort=user_name&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Username&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='useraccess.asp?sort=user_name&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth' width='20%'>Name</th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='useraccess.asp?sort=write_access&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Write Access&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='useraccess.asp?sort=write_access&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='useraccess.asp?sort=delete_access&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Delete Access&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='useraccess.asp?sort=delete_access&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='useraccess.asp?sort=disabled&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Disabled&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='useraccess.asp?sort=disabled&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	If access = "write" Or access = "delete" Then
		response.write "<th id='mediumth'>&nbsp;</th>"
	End If
	If access = "delete" Or recordid < 0 Then
 	  response.Write "<th id='mediumth'>&nbsp;</th>"
 	End If
	response.Write "</tr>"

	response.Write "<form action='adminaction.asp' id='form1' method='post' name='form1'>"

	response.write "<input type='hidden' name='RECORD' value='" & recordid & "'>"
	response.write "<input type='hidden' name='SORT' value='" & sortkey & "'>"
	response.write "<input type='hidden' name='DIRECTION' value='" & sortdir & "'>"

	'Read the form permission data.
	sqlString = "SELECT p.permission_id,p.role_id,r.role_name," & _
				"user_name,CONCAT(nickname,' ',last_name) AS name," & _
				"form_name,write_access,delete_access,disabled " & _
				"FROM (application_permissions p " & _
				"LEFT JOIN users u " & _
				"ON p.user_name=u.cwid) " & _
				"LEFT JOIN application_roles r " & _
				"ON p.role_id=r.role_id " & _
				"WHERE LOWER(p.application_name)='inspections' " & _
				"ORDER BY " & sortkey & " " & sortdir
	set rs = cn.Execute (sqlString)
	If Not rs.BOF Then
		rs.MoveFirst
	End If

	'If recordid<0, the user has selected "Add new record" so insert a blank data entry line
	'at the top of the form.
	If access = "write" Or access = "delete" Then
		If recordid < 0 Then
			Response.Write "<tr>"
			'Dropdown for form names.
			If session("err") = "form_name" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='form_name'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='form_name'>"
			End If
			Response.Write "<option value=''>"
			sqlString = "SELECT form_name FROM application_forms WHERE application_name='inspections'"
			Set rs3 = cn.Execute(sqlString)
			If Not rs3.BOF Then
				rs3.MoveFirst
				Do While Not rs3.EOF
					If UCase(Session("form_name")) = UCase(rs3("form_name")) Then
						response.write "<option value='" & rs3("form_name") & "' selected>" & rs3("form_name")
					Else
						response.write "<option value='" & rs3("form_name") & "'>" & rs3("form_name")
					End If
					rs3.MoveNext
				Loop
			End If
			rs3.Close
			response.write "</select>"
			response.write "</td>"

			'Dropdown for role names.
			If session("err") = "role_id" Then
				response.write "<td style='background-color:red;text-align:center'><select name='role_id'>"
			Else
				response.write "<td style='text-align:center'><select name='role_id'>"
			End If
			Response.Write "<option value=''>"
			strSQL = "SELECT role_id,role_name FROM application_roles WHERE application_name='inspections' ORDER BY role_name"
			Set rs3 = cn.Execute(strSQL)
			If Not rs3.BOF Then
				rs3.MoveFirst
				Do While Not rs3.EOF
					If Session("role_id") = rs3("role_id") Then
						response.write "<option value='" & rs3("role_id") & "' selected>" & rs3("role_name")
					Else
						response.write "<option value='" & rs3("role_id") & "'>" & rs3("role_name")
					End If
					rs3.MoveNext
				Loop
			End If
			rs3.Close
			response.write "</select>"
			response.write "</td>"

			If session("err") = "user_name" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='user_name'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='user_name'>"
			End If
			name = ""
			Response.Write "<option value=''>"
			If Session("allowed_user") = "everybody" Then
				Response.Write "<option value='everybody' selected>everybody"
			Else
				Response.Write "<option value='everybody'>everybody"
			End If
			sqlString = "SELECT cwid,CONCAT(nickname,' ',last_name) AS name " & _
						"FROM users " & _
						"WHERE status='Active' ORDER BY name"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Do While Not rs2.EOF
'					If UCase(session("user_name")) = UCase(rs2("unid")) Then
					If UCase(Session("user_name")) = UCase(rs2("cwid")) Then
'						response.write "<option value='" & rs2("unid") & "' selected>" & rs2("unid") & " - " & rs2("first_name") & " " & rs2("last_name")
						Response.Write "<option value='" & UCase(rs2("cwid")) & "' selected>" & UCase(rs2("cwid")) & " - " & rs2("name")
'						name = rs2("first_name") & " " & rs2("last_name")
					Else
'						response.write "<option value='" & rs2("unid") & "'>" & rs2("unid") & " - " & rs2("first_name") & " " & rs2("last_name")
						Response.Write "<option value='" & UCase(rs2("cwid")) & "'>" & UCase(rs2("cwid")) & " - " & rs2("name")
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			Response.Write "</select></td>"
			
			If Session("name") <> "" Then
				Response.Write "<td id='mediumtd' style='text-align:center'>" & Session("name") & "</td>"
			Else
				Response.Write "<td id='mediumtd' style='text-align:center'>&nbsp;</td>"
			End If

			If Session("write_access") <> "" Then
				response.write "<td id='mediumtd' style='text-align:center'><input type='checkbox' name='write_access' checked />"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><input type='checkbox' name='write_access' />"
			End If

			If Session("delete_access") <> "" Then
				response.write "<td id='mediumtd' style='text-align:center'><input type='checkbox' name='delete_access' checked />"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><input type='checkbox' name='delete_access' />"
			End If

			If Session("disabled") <> "" Then
				response.write "<td id='mediumtd' style='text-align:center'><input type='checkbox' name='disabled' checked />"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><input type='checkbox' name='disabled' />"
			End If

			response.write "<td id='mediumtd' style='text-align:center'><input type='submit' value='Submit' id='submit1' name='submit1' style='font-size:8pt;background-color:white' title='Apply changes to this record' /></td>"

			Response.Write "<td id='mediumtd' style='text-align:center'><a href='useraccess.asp?clear=true&sort=" & sortkey & "&direction=" & sortdir & "' title='Cancel changes to this record'>Cancel</a></td>"

			Response.Write "</tr>"
		End If
	End If

	Do While Not rs.EOF
		If CInt(rs("permission_id")) = CInt(recordid) Then
			'Draw the data entry line
			Response.Write "<tr>"
			'Dropdown for form names.
			If session("err") = "form_name" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='form_name'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='form_name'>"
			End If
			Response.Write "<option value=''>"
			sqlString = "SELECT form_name FROM application_forms WHERE application_name='inspections'"
			Set rs3 = cn.Execute(sqlString)
			If Not rs3.BOF Then
				rs3.MoveFirst
				Do While Not rs3.EOF
					If UCase(rs("form_name")) = UCase(rs3("form_name")) Then
						response.write "<option value='" & rs3("form_name") & "' selected>" & rs3("form_name")
					Else
						response.write "<option value='" & rs3("form_name") & "'>" & rs3("form_name")
					End If
					rs3.MoveNext
				Loop
			End If
			rs3.Close
			response.write "</select>"
			response.write "</td>"

			'Dropdown for role names.
			If session("err") = "role_id" Then
				response.write "<td style='background-color:red;text-align:center'><select name='role_id'>"
			Else
				response.write "<td style='text-align:center'><select name='role_id'>"
			End If
			Response.Write "<option value=''>"
			strSQL = "SELECT role_id,role_name FROM application_roles WHERE application_name='inspections' ORDER BY role_name"
			Set rs3 = cn.Execute(strSQL)
			If Not rs3.BOF Then
				rs3.MoveFirst
				Do While Not rs3.EOF
					If Session("role_id") = rs3("role_id") Then
						response.write "<option value='" & rs3("role_id") & "' selected>" & rs3("role_name")
					ElseIf rs("role_id") = rs3("role_id") Then
						response.write "<option value='" & rs3("role_id") & "' selected>" & rs3("role_name")
					Else
						response.write "<option value='" & rs3("role_id") & "'>" & rs3("role_name")
					End If
					rs3.MoveNext
				Loop
			End If
			rs3.Close
			response.write "</select>"
			response.write "</td>"

			If session("err") = "user_name" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='user_name'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='user_name'>"
			End If
			name = ""
			Response.Write "<option value=''>"
			If rs("user_name") = "everybody" Then
				Response.Write "<option value='everybody' selected>everybody"
			Else
				Response.Write "<option value='everybody'>everybody"
			End If
			sqlString = "SELECT cwid,CONCAT(nickname,' ',last_name) AS name " & _
						"FROM users " & _
						"WHERE status='Active' ORDER BY name"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Do While Not rs2.EOF
'					If UCase(rs("user_name")) = UCase(rs2("unid")) Then
					If UCase(rs("user_name")) = UCase(rs2("cwid")) Then
'						response.write "<option value='" & rs2("unid") & "' selected>" & rs2("unid") & " - " & rs2("first_name") & " " & rs2("last_name")
						Response.Write "<option value='" & UCase(rs2("cwid")) & "' selected>" & UCase(rs2("cwid")) & " - " & rs2("name")
'						name = rs2("first_name") & " " & rs2("last_name")
					Else
'						response.write "<option value='" & rs2("unid") & "'>" & rs2("unid") & " - " & rs2("first_name") & " " & rs2("last_name")
						Response.Write "<option value='" & UCase(rs2("cwid")) & "'>" & UCase(rs2("cwid")) & " - " & rs2("name")
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			Response.Write "</select></td>"

			If rs("name") <> "" Then
				Response.Write "<td id='mediumtd' style='text-align:center'>" & rs("name") & "</td>"
			Else
				Response.Write "<td id='mediumtd' style='text-align:center'>&nbsp;</td>"
			End If

			If rs("write_access") <> 0 Then
				response.write "<td id='mediumtd' style='text-align:center'><input type='checkbox' name='write_access' checked />"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><input type='checkbox' name='write_access' />"
			End If

			If rs("delete_access") <> 0 Then
				response.write "<td id='mediumtd' style='text-align:center'><input type='checkbox' name='delete_access' checked />"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><input type='checkbox' name='delete_access' />"
			End If

			If rs("disabled") <> 0 Then
				response.write "<td id='mediumtd' style='text-align:center'><input type='checkbox' name='disabled' checked />"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><input type='checkbox' name='disabled' />"
			End If

			If access = "write" Or access = "delete" Then
				response.write "<td id='mediumtd' style='text-align:center'><input type='submit' value='Submit' id='submit1' name='submit1' style='font-size:8pt;background-color:white' title='Apply changes for this record' /></td>"
			End If

'			If access = "delete" Then
'				recid = rs("permission_id")
'				response.write "<td id='mediumtd'><a href='javascript:doDelete(" & recid & ");' title='Delete this record'>Delete</a></td>"
'			End If
			Response.Write "<td id='mediumtd' style='text-align:center'><a href='useraccess.asp?clear=true&sort=" & sortkey & "&direction=" & sortdir & "' title='Cancel changes to this record'>Cancel</a></td>"

			Response.Write "</tr>"

		Else
			'Draw the history records
			response.write "<tr>"
			If rs("form_name") <> "" Then
				response.write "<td id='mediumtd' style='text-align:center'>" & rs("form_name") & "</td>"
			Else
				Response.Write "<td id='mediumtd' style='text-align:center'>&nbsp;</td>"
			End If

			If Not IsNull(rs("role_id")) Then
				response.write "<td id='mediumtd' style='text-align:center'>" & rs("role_name") & "</td>"
			Else
				Response.Write "<td id='mediumtd' style='text-align:center'>&nbsp;</td>"
			End If

			If rs("user_name") <> "" Then
				response.write "<td id='mediumtd' style='text-align:center'>" & rs("user_name") & "</td>"
			Else
				Response.Write "<td id='mediumtd' style='text-align:center'>&nbsp;</td>"
			End If
	    
			If rs("name") <> "" Then
				Response.Write "<td id='mediumtd' style='text-align:center'>" & rs("name") & "</td>"
			Else
				Response.Write "<td id='mediumtd' style='text-align:center'>&nbsp;</td>"
			End If

			If rs("write_access") <> 0 Then
				response.write "<td id='mediumtd' style='text-align:center'><input type='checkbox' checked disabled /></td>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><input type='checkbox' disabled /></td>"
			End If

			If rs("delete_access") <> 0 Then
				response.write "<td id='mediumtd' style='text-align:center'><input type='checkbox' checked disabled /></td>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><input type='checkbox' disabled /></td>"
			End If

			If rs("disabled") <> 0 Then
				response.write "<td id='mediumtd' style='text-align:center'><input type='checkbox' checked disabled /></td>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><input type='checkbox' disabled /></td>"
			End If

			If access = "write" Or access = "delete" Then
				response.write "<td id='mediumtd' style='text-align:center'><a href='useraccess.asp?record_id=" & rs("permission_id") & "&sort=" & sortkey & "&direction=" & sortdir & "' title='Edit this record'>Edit</a></td>"
			End If

			If access = "delete" Then
				recid = rs("permission_id")
				response.write "<td id='mediumtd'><a href='javascript:doDelete(" & recid & ");' title='Delete this record'>Delete</a></td>"
			End If
			response.write "</tr>"

		End If
		rs.Movenext
	loop
	rs.Close
'	rs2.Close

	Response.Write "</form>"
	Response.Write "</table>"
	Response.Write "</body>"

	Set rs = Nothing
	Set rs2 = Nothing
	Set rs3 = Nothing
	cn.Close
	Set cn = Nothing

Else
	response.write "<h1>You don't have permission to access this page.</h1>"
	response.write "<br />"
	response.write "<a href='" & request("http_referer") & "'>Return to previous page</a>"
End If
%>
</html>
