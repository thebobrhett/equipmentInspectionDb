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
</script>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Maintain Inspections Role Members</title>
<link rel=STYLESHEET href='formstyle.css' type='text/css'>
<style>
<!--
  a     { text-decoration:underline; }
-->
</style> 
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="InspectionsFunctions.asp"-->
</head>

<%
'*************
' Revision History
' 
' Keith Brooks - Monday, August 1, 2011
'   Creation.
'*************

dim objConn
dim objRS
Dim objRS2
Dim objRS3
Dim strSQL
dim recordid
Dim currentuser
Dim access
Dim name
Dim recid

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("inspections","rolemembers",currentuser)
If access <> "none" Then

	set objConn = CreateObject("adodb.connection")
	objConn.Open = "driver={MySQL ODBC 3.51 Driver};option=16387;server=richmond.aksa.local;User=assetmgtuser;password=asset;DATABASE=asset_management;"
	set objRS = CreateObject("adodb.recordset")
	set objRS3 = CreateObject("adodb.recordset")
	set objRS2 = CreateObject("adodb.recordset")

	If session("err") <> "" And session("err") <> "NONE" Then
		Response.Write "<body ID='beigebody' onload='document.form1." & session("err") & ".focus();'>"
	ElseIf request("recordnum") <> "" Then
		Response.Write "<body ID='beigebody' onload='document.form1.form_name.focus();'>"
		session("focus") = "NONE"
	Else
		Response.Write "<body ID='beigebody'>"
	End If

	If request("recordnum") <> "" Then
		If IsNumeric(request("recordnum")) Then
			recordid = request("recordnum")
		Else
			recordid = 0
		End If
	Else
		recordid = 0
	End If

	%>
		<div id="PleaseWait" style="display: none; text-align:center; color:White; vertical-align:top; position:absolute; top:5px; left:5px; z-index:100">
			<table id="MyTable" style="background-color:blue">
				<tr><td style="width:95px; text-size:12px; color:white; font-weight:bold">Please Wait...</td></tr>
			</table>
		</div>
	<%

	response.write "<table ID='headertable' width='100%'>"
	response.write "<tr>"
	response.write "<td ID='headertd' style='text-align:left' width='20%'><a href='adminmenu.asp' title='Open the administration main menu'>Menu</a>"
	response.write "<td ID='headertd' style='text-align:center' width='60%'><h1/>Maintain Inspections Role Members</td>"
	response.write "<td ID='headertd' style='text-align:right' width='20%'>&nbsp;</td>"
	response.write "</tr>"

	response.write "<tr>"
	response.write "<td id='headertd'>&nbsp;</td>"
	If access = "write" Or access = "delete" Then
		response.write "<td ID='headertd' style='text-align:center;font-size:12px'><a href='rolemembers.asp?recordnum=-1' title='Add a new role member record'>Add new record</a></td>"
	Else
		Response.Write "<td id='headertd'>&nbsp;</td>"
	End If
	response.write "<td id='headertd'>&nbsp;</td>"
	response.write "</tr>"
	response.write "</table>"

	'Draw the header
	response.Write "<table align='center' width='90%' border='1' style='font-size:12px'>"
	response.Write "<tr>"
	response.write "<th id='mediumth'>Role</th>"
	response.Write "<th id='mediumth'>Username</th>"
	response.Write "<th id='mediumth' width='20%'>Name</th>"
	If access = "write" Or access = "delete" Then
		response.write "<th id='mediumth'>&nbsp;</th>"
	End If
	If access = "delete" Or recordid < 0 Then
 	  response.Write "<th id='mediumth'>&nbsp;</th>"
 	End If
	response.Write "</tr>"

	response.Write "<form action='adminaction.asp' id='form1' method='post' name='form1'>"

	response.write "<input type='hidden' name='RECORD' value='" & recordid & "'>"

	'Read the role member data.
	strSQL = "SELECT p.role_member_id,p.application_role_id,r.role_name," & _
			"user_name,CONCAT(nickname,' ',last_name) AS name " & _
			"FROM (application_role_members p LEFT JOIN users u " & _
			"ON p.user_name=u.cwid) " & _
			"LEFT JOIN application_roles r " & _
			"ON p.application_role_id=r.role_id " & _
			"WHERE r.application_name='inspections' " & _
			"ORDER BY role_name,user_name"
	set objRS = objConn.Execute (strSQL)
	If Not objRS.BOF Then
		objRs.MoveFirst
	End If

	'If recordid<0, the user has selected "Add new record" so insert a blank data entry line
	'at the top of the form.
	If access = "write" Or access = "delete" Then
		If recordid < 0 Then
			Response.Write "<tr>"
			'Dropdown for role names.
			If session("err") = "application_role_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='application_role_id'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='application_role_id'>"
			End If
			Response.Write "<option value=''>"
			strSQL = "SELECT role_id,role_name FROM application_roles " & _
						"WHERE application_name='inspections' ORDER BY role_name"
			Set objRS3 = objConn.Execute(strSQL)
			If Not objRS3.BOF Then
				objRS3.MoveFirst
				Do While Not objRS3.EOF
					If Session("application_role_id") = objRS3("role_id") Then
						response.write "<option value='" & objRS3("role_id") & "' selected>" & objRS3("role_name")
					Else
						response.write "<option value='" & objRS3("role_id") & "'>" & objRS3("role_name")
					End If
					objRS3.MoveNext
				Loop
			End If
			objRS3.Close
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
			strSQL = "SELECT cwid,CONCAT(nickname,' ',last_name) AS name " & _
					"FROM users " & _
					"WHERE status='Active' ORDER BY name"
			Set objRS2 = objConn.Execute(strSQL)
			If Not objRS2.BOF Then
				Do While Not objRS2.EOF
					If UCase(Session("user_name")) = UCase(objRS2("cwid")) Then
						Response.Write "<option value='" & UCase(objRS2("cwid")) & "' selected>" & UCase(objRS2("cwid")) & " - " & objRS2("name")
					Else
						Response.Write "<option value='" & UCase(objRS2("cwid")) & "'>" & UCase(objRS2("cwid")) & " - " & objRS2("name")
					End If
					objRS2.MoveNext
				Loop
			End If
			objRS2.Close
			Response.Write "</select></td>"
			
			If Session("name") <> "" Then
				Response.Write "<td id='mediumtd' style='text-align:center'>" & Session("name") & "</td>"
			Else
				Response.Write "<td id='mediumtd' style='text-align:center'>&nbsp;</td>"
			End If

			response.write "<td id='mediumtd' style='text-align:center'><input type='submit' value='Submit' id='submit1' name='submit1' title='Apply changes to this record' /></td>"

			Response.Write "<td id='mediumtd' style='text-align:center'><a href='rolemembers.asp' title='Cancel changes to this record'>Cancel</a></td>"

			Response.Write "</tr>"
		End If
	End If

	Do While Not objRS.EOF
		If CInt(objRS("role_member_id")) = CInt(recordid) Then
			'Draw the data entry line
			Response.Write "<tr>"

			'Dropdown for role names.
			If session("err") = "application_role_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='application_role_id'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='application_role_id'>"
			End If
			Response.Write "<option value=''>"
			strSQL = "SELECT role_id,role_name FROM application_roles " & _
						"WHERE application_name='inspections' ORDER BY role_name"
			Set objRS3 = objConn.Execute(strSQL)
			If Not objRS3.BOF Then
				objRS3.MoveFirst
				Do While Not objRS3.EOF
					If Session("application_role_id") = objRS3("role_id") Then
						response.write "<option value='" & objRS3("role_id") & "' selected>" & objRS3("role_name")
					ElseIf objRS("application_role_id") = objRS3("role_id") Then
						response.write "<option value='" & objRS3("role_id") & "' selected>" & objRS3("role_name")
					Else
						response.write "<option value='" & objRS3("role_id") & "'>" & objRS3("role_name")
					End If
					objRS3.MoveNext
				Loop
			End If
			objRS3.Close
			response.write "</select>"
			response.write "</td>"

			If session("err") = "user_name" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='user_name'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='user_name'>"
			End If
			name = ""
			Response.Write "<option value=''>"
			If objRS("user_name") = "everybody" Then
				Response.Write "<option value='everybody' selected>everybody"
			Else
				Response.Write "<option value='everybody'>everybody"
			End If
			strSQL = "SELECT cwid,CONCAT(nickname,' ',last_name) AS name " & _
					"FROM users " & _
					"WHERE status='Active' ORDER BY name"
			Set objRS2 = objConn.Execute(strSQL)
			If Not objRS2.BOF Then
				objRS2.MoveFirst
				Do While Not objRS2.EOF
					If UCase(objRS("user_name")) = UCase(objRS2("cwid")) Then
						Response.Write "<option value='" & UCase(objRS2("cwid")) & "' selected>" & UCase(objRS2("cwid")) & " - " & objRS2("name")
					Else
						Response.Write "<option value='" & UCase(objRS2("cwid")) & "'>" & UCase(objRS2("cwid")) & " - " & objRS2("name")
					End If
					objRS2.MoveNext
				Loop
			End If
			objRS2.Close
			Response.Write "</select></td>"

			If objRS("name") <> "" Then
				Response.Write "<td id='mediumtd' style='text-align:center'>" & objRS("name") & "</td>"
			Else
				Response.Write "<td id='mediumtd' style='text-align:center'>&nbsp;</td>"
			End If

			If access = "write" Or access = "delete" Then
				response.write "<td id='mediumtd' style='text-align:center'><input type='submit' value='Submit' id='submit1' name='submit1' title='Apply changes for this record' /></td>"
			End If

			If access = "delete" Then
				recid = objRS("role_member_id")
				response.write "<td id='mediumtd'><a href='javascript:doDelete(" & recid & ");' title='Delete this record'>Delete</a></td>"
			End If

			Response.Write "</tr>"

		Else
			'Draw the history records
			response.write "<tr>"
			If Not IsNull(objRS("application_role_id")) Then
				response.write "<td id='mediumtd' style='text-align:center'>" & objRS("role_name") & "</td>"
			Else
				Response.Write "<td id='mediumtd' style='text-align:center'>&nbsp;</td>"
			End If

			If objRS("user_name") <> "" Then
				response.write "<td id='mediumtd' style='text-align:center'>" & objRS("user_name") & "</td>"
			Else
				Response.Write "<td id='mediumtd' style='text-align:center'>&nbsp;</td>"
			End If
	    
			If objRS("name") <> "" Then
				Response.Write "<td id='mediumtd' style='text-align:center'>" & objRS("name") & "</td>"
			Else
				Response.Write "<td id='mediumtd' style='text-align:center'>&nbsp;</td>"
			End If

			If access = "write" Or access = "delete" Then
				response.write "<td id='mediumtd' style='text-align:center'><a href='rolemembers.asp?recordnum=" & objRS("role_member_id") & "' title='Edit this record'>Edit</a></td>"
			End If

			If access = "delete" Then
				recid = objRS("role_member_id")
				response.write "<td id='mediumtd' style='text-align:center'><a href='javascript:doDelete(" & recid & ");' title='Delete this record'>Delete</a></td>"
			End If
			response.write "</tr>"

		End If
		objRS.Movenext
	loop
	objRS.Close
'	objRS2.Close

	Set objRS = Nothing
	Set objRS2 = Nothing
	Set objRS3 = Nothing
	objConn.Close
	Set objConn = Nothing
	
	Response.Write "</form>"
	Response.Write "</table>"
	Response.Write "</body>"

Else
	response.write "<h1>You don't have permission to access this page.</h1>"
	response.write "<br />"
	response.write "<a href='" & request("http_referer") & "'>Return to previous page</a>"
End If
%>
</html>
