<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv='Expires' content='0'></meta>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'></meta>
<title>Inspections Admin Action</title>
<link rel=STYLESHEET href='formstyle.css' type='text/css'>
<!--#include file="InspectionsFunctions.asp"-->
<%
Function GetInsertFields(fieldNames,fieldVals)
	Dim insertFields
	Dim count
	insertFields = ""
	For count = 0 To UBound(fieldNames)
		If fieldVals(count) <> "" Then
			If insertFields = "" Then
				insertFields = fieldNames(count)
			Else
				insertFields = insertFields & "," & fieldNames(count)
			End If
		End If
	Next
	GetInsertFields = insertFields
End Function

Function GetInsertVals(fieldTypes,fieldVals)
	Dim insertVals
	Dim count
	insertVals = ""
	For count = 0 To UBound(fieldNames)
		If insertVals = "" Then
			If fieldVals(count) <> "" Then
				If fieldTypes(count) = "Text" Then
					insertVals = "'" & fieldVals(count) & "'"
				ElseIf fieldTypes(count) = "Date" Then
					insertVals = "'" & FormatMySQLDateTime(fieldVals(count)) & "'"
				Else
					insertVals = fieldVals(count)
				End If
			End If
		Else
			If fieldVals(count) <> "" Then
				If fieldTypes(count) = "Text" Then
					insertVals = insertVals & ",'" & fieldVals(count) & "'"
				ElseIf fieldTypes(count) = "Date" Then
					insertVals = insertVals & ",'" & FormatMySQLDateTime(fieldVals(count)) & "'"
				Else
					insertVals = insertVals & "," & fieldVals(count)
				End If
			End If
		End If
	Next
	GetInsertVals = insertVals
End Function

Function GetUpdateString(fieldNames,fieldVals,fieldTypes)
	Dim updateString
	Dim count
	updateString = ""
	For count = 0 To UBound(fieldNames)
		If updateString = "" Then
			If fieldVals(count) <> "" Then
				If fieldTypes(count) = "Text" Then
					updateString = fieldNames(count) & "='" & fieldVals(count) & "'"
				ElseIf fieldTypes(count) = "Date" Then
					updateString = fieldNames(count) & "='" & FormatMySQLDateTime(fieldVals(count)) & "'"
				Else
					updateString = fieldNames(count) & "=" & fieldVals(count)
				End If
			Else
				updateString = fieldNames(count) & "=null"
			End If
		Else
			If fieldVals(count) <> "" Then
				If fieldTypes(count) = "Text" Then
					updateString = updateString & "," & fieldNames(count) & "='" & fieldVals(count) & "'"
				ElseIf fieldTypes(count) = "Date" Then
					updateString = updateString & "," & fieldNames(count) & "='" & FormatMySQLDateTime(fieldVals(count)) & "'"
				Else
					updateString = updateString & "," & fieldNames(count) & "=" & fieldVals(count)
				End If
			Else
				updateString = updateString & "," & fieldNames(count) & "=null"
			End If
		End If
	Next
	GetUpdateString = updateString
End Function

Function WriteDeleteQuery(tableName,idName)
	If IsNumeric(Request("RECORD")) Then
		WriteDeleteQuery = "DELETE FROM " & tableName & " WHERE " & idName & "=" & Request("RECORD")
	Else
		WriteDeleteQuery = "DELETE FROM " & tableName & " WHERE " & idName & "='" & Request("RECORD") & "'"
	End If
End Function

Function WriteInsertQuery(tableName,insertFields,insertVals)
	WriteInsertQuery = "INSERT INTO " & tableName & " (" & insertFields & ") VALUES (" & insertVals & ")"
End Function

Function WriteSelectQuery(fieldName,tableName,idName)
	If IsNumeric(Request("RECORD")) Then
		WriteSelectQuery = "SELECT " & fieldName & " FROM " & tableName & " WHERE " & idName & "=" & Request("RECORD")
	Else
		WriteSelectQuery = "SELECT " & fieldName & " FROM " & tableName & " WHERE " & idName & "='" & Request("RECORD") & "'"
	End If
End Function

Function WriteUpdateQuery(tableName,updateString,idName)
	If IsNumeric(Request("RECORD")) Then
		WriteUpdateQuery = "UPDATE " & tableName & " SET " & updateString & " WHERE " & idName & "=" & Request("RECORD")
	Else
		WriteUpdateQuery = "UPDATE " & tableName & " SET " & updateString & " WHERE " & idName & "='" & Request("RECORD") & "'"
	End If
End Function

Function WriteAuditInsertQuery(user,idVal,tableName,fieldName,oldVal,newVal,auditType)
	WriteAuditInsertQuery = "INSERT INTO admin_audit_trail (change_user,change_table_id,change_table,change_field,old_value,new_value,change_type) " & _
					"VALUES ('" & user & "','" & idVal & "','" & tableName & "','" & fieldName & "','" & oldVal & "','" & newVal & "','" & auditType & "')"
End Function

Function ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes,ByRef tableID)
	Dim cn
	Dim rs
	Dim status
	On Error Resume Next
	status = True
	set cn = CreateObject("adodb.connection")
	cn.Open = DBString
	If Err.number <> 0 Then
		status = False
		Exit Function
	End If
	set rs = CreateObject("adodb.recordset")
	If Request.QueryString("action") = "delete" Then
		'Delete the record.
		Set rs = cn.Execute(WriteDeleteQuery(tableName,idName))
		If Err.number <> 0 Then
			Session("err") = "db"
			Exit Function
		End If
		'Write the change to the audit trail table.
		Set rs = cn.Execute(WriteAuditInsertQuery(currentuser,Request("RECORD"),tableName,idName,Request("RECORD"),"null","delete"))
		If Err.number <> 0 Then
			status = False
			Exit Function
		End If
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = ""
		Next
	Else
		'If the record number < 0, insert a record; otherwise, update the specified record.
		If request("RECORD") <> "" Then
			If request("RECORD") = -1 Or Request("RECORD") = "-1" Then
'				If 1=2 Then
				Set rs = cn.Execute(WriteInsertQuery(tableName,GetInsertFields(fieldNames,fieldVals),GetInsertVals(fieldTypes,fieldVals)))
				If Err.number <> 0 Then
					status = False
					Exit Function
				End If
				'Get the id number that was just assigned.
				sqlString = "SELECT LAST_INSERT_ID()"
				Set rs = cn.Execute(sqlString)
				If Err.number <> 0 Then
					status = False
					Exit Function
				End If
				If Not rs.BOF Then
					tableID = rs(0)
				Else
					tableID = 0
				End If
				rs.Close
				'If the ID is not an autoincrement, use the first field value.
				If tableID = 0 Then
					tableID = fieldVals(0)
				End If
				'Write the changes to the audit trail.
				Set rs = cn.Execute(WriteAuditInsertQuery(currentuser,tableID,tableName,idName,"null",tableID,"insert"))
				For count = 0 To UBound(fieldNames)
					Set rs = cn.Execute(WriteAuditInsertQuery(currentuser,tableID,tableName,fieldNames(count),"null",fieldVals(count),"insert"))
					If Err.number <> 0 Then
						status = False
						Exit Function
					End If
				Next
'				Else
'					Response.Write "sqlString = " & WriteInsertQuery(tableName,GetInsertFields(fieldNames,fieldVals),GetInsertVals(fieldTypes,fieldVals))
'				End If
			Else
				'Get the existing field values for the audit trail.
				ReDim oldVals(UBound(fieldNames))
				For count = 0 To UBound(fieldNames)
					Set rs = cn.Execute(WriteSelectQuery(fieldNames(count),tableName,idName))
					If Err.number <> 0 Then
						status = False
						Exit Function
					End If
					If Not rs.BOF Then
						rs.MoveFirst
						If Not IsNull(rs(0)) Then
							oldVals(count) = rs(0)
						Else
							oldVals(count) = "null"
						End If
					Else
						oldVals(count) = "null"
					End If
					rs.Close
				Next
				
				'Update the record.
'				If 1=2 Then
				Set rs = cn.Execute(WriteUpdateQuery(tableName,GetUpdateString(fieldNames,fieldVals,fieldTypes),idName))
				If Err.number <> 0 Then
					status = False
					Exit Function
				End If
'				Else
'					Response.Write "SQL = " & WriteUpdateQuery(tableName,GetUpdateString(fieldNames,fieldVals,fieldTypes),idName)
'				End If

				'Write the changes to the audit trail table.
				For count = 0 To UBound(fieldNames)
					Set rs = cn.Execute(WriteAuditInsertQuery(currentuser,Request("RECORD"),tableName,fieldNames(count),oldVals(count),fieldVals(count),"update"))
					If Err.number <> 0 Then
						status = False
						Exit Function
					End If
				Next
			End If
				
		End If
		Session.Contents.RemoveAll
'		For count = 0 To UBound(fieldNames)
'			Session(fieldNames(count)) = ""
'		Next
	End If
	Set rs = Nothing
	cn.Close
	Set cn = Nothing
	On Error Goto 0
	ProcessChange = status
End Function

Function ProcessUserChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
	Dim cn
	Dim cn2
	Dim rs
	Dim rs2
	Dim status
	On Error Resume Next
	status = True
	set cn = CreateObject("adodb.connection")
	cn.Open = "driver={MySQL ODBC 3.51 Driver};option=16387;server=richmond.aksa.local;User=assetmgtuser;password=asset;DATABASE=asset_management;"
	If Err.number <> 0 Then
		status = False
		Exit Function
	End If
	set rs = CreateObject("adodb.recordset")
	set cn2 = CreateObject("adodb.connection")
	cn2.Open = DBString
	If Err.number <> 0 Then
		status = False
		Exit Function
	End If
	set rs2 = CreateObject("adodb.recordset")
	If Request.QueryString("action") = "delete" Then
		'Delete the record.
		Set rs = cn.Execute(WriteDeleteQuery(tableName,idName))
		If Err.number <> 0 Then
			Session("err") = "db"
			Exit Function
		End If
		'Write the change to the audit trail table.
		Set rs2 = cn2.Execute(WriteAuditInsertQuery(currentuser,Request("RECORD"),tableName,idName,Request("RECORD"),"null","delete"))
		If Err.number <> 0 Then
			status = False
			Exit Function
		End If
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = ""
		Next
	Else
		'If the record number < 0, insert a record; otherwise, update the specified record.
		If request("RECORD") <> "" Then
			If request("RECORD") = -1 Or Request("RECORD") = "-1" Then
				Set rs = cn.Execute(WriteInsertQuery(tableName,GetInsertFields(fieldNames,fieldVals),GetInsertVals(fieldTypes,fieldVals)))
				If Err.number <> 0 Then
					status = False
					Exit Function
				End If
				'Get the id number that was just assigned.
				sqlString = "SELECT LAST_INSERT_ID()"
				Set rs = cn.Execute(sqlString)
				If Err.number <> 0 Then
					status = False
					Exit Function
				End If
				If Not rs.BOF Then
					tableID = rs(0)
				Else
					tableID = 0
				End If
				rs.Close
				'If the ID is not an autoincrement, use the first field value.
				If tableID = 0 Then
					tableID = fieldVals(0)
				End If
				'Write the changes to the audit trail.
				Set rs2 = cn2.Execute(WriteAuditInsertQuery(currentuser,tableID,tableName,idName,"null",tableID,"insert"))
				For count = 0 To UBound(fieldNames)
					Set rs2 = cn2.Execute(WriteAuditInsertQuery(currentuser,tableID,tableName,fieldNames(count),"null",fieldVals(count),"insert"))
					If Err.number <> 0 Then
						status = False
						Exit Function
					End If
				Next
			Else
				'Get the existing field values for the audit trail.
				ReDim oldVals(UBound(fieldNames))
				For count = 0 To UBound(fieldNames)
					Set rs = cn.Execute(WriteSelectQuery(fieldNames(count),tableName,idName))
					If Err.number <> 0 Then
						status = False
						Exit Function
					End If
					If Not rs.BOF Then
						rs.MoveFirst
						If Not IsNull(rs(0)) Then
							oldVals(count) = rs(0)
						Else
							oldVals(count) = "null"
						End If
					Else
						oldVals(count) = "null"
					End If
					rs.Close
				Next
				
				'Update the record.
				Set rs = cn.Execute(WriteUpdateQuery(tableName,GetUpdateString(fieldNames,fieldVals,fieldTypes),idName))
				If Err.number <> 0 Then
					status = False
					Exit Function
				End If

				'Write the changes to the audit trail table.
				For count = 0 To UBound(fieldNames)
					Set rs2 = cn2.Execute(WriteAuditInsertQuery(currentuser,Request("RECORD"),tableName,fieldNames(count),oldVals(count),fieldVals(count),"update"))
					If Err.number <> 0 Then
						status = False
						Exit Function
					End If
				Next
			End If
				
		End If
		Session.Contents.RemoveAll
'		For count = 1 To UBound(fieldNames)
'			Session(fieldNames(count)) = ""
'		Next
	End If
	Set rs = Nothing
	Set rs2 = Nothing
	cn.Close
	cn2.Close
	Set cn = Nothing
	Set cn2 = Nothing
	On Error Goto 0
	ProcessUserChange = status
End Function
%>
</head>

<%
'*************
' Revision History
' 
' Keith Brooks - Tuesday, February 15, 2011
'   Creation.
'*************

'on error resume next
Dim fieldVals()
Dim fieldNames()
Dim fieldTypes()
Dim oldVals()
Dim oldVal
Dim sqlString
Dim reload
Dim samePage
Dim NewID
Dim currentuser
Dim count
Dim insertFields
Dim insertVals
Dim updateString
Dim tableID
Dim tableName
Dim idName
Dim status

reload = "NONE"
session("err") = "NONE"
samePage = False
NewID = 0

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

tableID = 0

If InStr(request("http_referer"),"useraccess") > 0 Then
	tableName = "application_permissions"
	idName = "permission_id"
	ReDim fieldNames(6)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "application_name"
	fieldNames(1) = "user_name"
	fieldNames(2) = "form_name"
	fieldNames(3) = "role_id"
	fieldNames(4) = "write_access"
	fieldNames(5) = "delete_access"
	fieldNames(6) = "disabled"
	For count = 0 To UBound(fieldNames)
		If fieldNames(count) = "application_name" Then
			fieldVals(count) = "inspections"
		ElseIf fieldNames(count) = "user_name" Or fieldNames(count) = "form_name" Or fieldNames(count) = "role_id" Then
			fieldVals(count) = Request(fieldNames(count))
		Else
			If Request(fieldNames(count)) <> "" Then
				fieldVals(count) = 1
			Else
				fieldVals(count) = 0
			End If
		End If
	Next
	fieldTypes(0) = "Text"
	fieldTypes(1) = "Text"
	fieldTypes(2) = "Text"
	fieldTypes(3) = "Number"
	fieldTypes(4) = "Number"
	fieldTypes(5) = "Number"
	fieldTypes(6) = "Number"

	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("form_name") = "" Then
			session("err") = "form_name"
			session("saveval") = request("form_name")
		ElseIf Request("role_id") = "" And request("user_name") = "" Then
			session("err") = "role_id"
			session("saveval") = request("role_id")
		End If
	End If

	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessUserChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 1 To UBound(fieldNames)
				If count > 2 And fieldVals(count) = "0" Then
'					Session(fieldNames(count)) = ""
					Session.Contents.Remove(FieldNames(count))
				Else
					Session(fieldNames(count)) = fieldVals(count)
				End If
			Next
		End If
	Else
		For count = 1 To UBound(fieldNames)
			If count > 2 And fieldVals(count) = "0" Then
'				Session(fieldNames(count)) = ""
				Session.Contents.Remove(FieldNames(count))
			Else
				Session(fieldNames(count)) = fieldVals(count)
			End If
		Next
	End If
	
ElseIf InStr(request("http_referer"),"rolemembers") > 0 Then
	tableName = "application_role_members"
	idName = "role_member_id"
	ReDim fieldNames(1)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "application_role_id"
	fieldNames(1) = "user_name"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Number"
	fieldTypes(1) = "Text"

	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("application_role_id") = "" Then
			session("err") = "application_role_id"
			session("saveval") = request("application_role_id")
		ElseIf Request("user_name") = "" Then
			session("err") = "user_name"
			session("saveval") = request("user_name")
		End If
	End If

	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessUserChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 1 To UBound(fieldNames)
				If count > 2 And fieldVals(count) = "0" Then
'					Session(fieldNames(count)) = ""
					Session.Contents.Remove(FieldNames(count))
				Else
					Session(fieldNames(count)) = fieldVals(count)
				End If
			Next
		End If
	Else
		For count = 1 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If
	
ElseIf InStr(request("http_referer"),"equipmentitems") > 0 Then
	tableName = "equipment_items"
	idName = "equipment_item_id"
	ReDim fieldNames(6)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "equipment_item_name"
	fieldNames(1) = "equipment_item_tag"
	fieldNames(2) = "equipment_item_description"
	fieldNames(3) = "equipment_type_id"
	fieldNames(4) = "assembly"
	fieldNames(5) = "area"
	fieldNames(6) = "conservation_vent"
	fieldTypes(0) = "Text"
	fieldTypes(1) = "Text"
	fieldTypes(2) = "Text"
	fieldTypes(3) = "Number"
	fieldTypes(4) = "Text"
	fieldTypes(5) = "Text"
	fieldTypes(6) = "Number"
	For count = 0 To UBound(fieldNames)
		'Add special case for checkbox field.
		If count = 6 Then
			If Request(fieldNames(count)) = "" Then
				fieldVals(count) = 0
			Else
				fieldVals(count) = 1
			End If
		Else
			fieldVals(count) = Request(fieldNames(count))
		End If
	Next

	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("equipment_item_name") = "" Then
			session("err") = "equipment_item_name"
			session("saveval") = request("equipment_item_name")
		ElseIf request("equipment_item_tag") = "" Then
			session("err") = "equipment_item_tag"
			session("saveval") = request("equipment_item_tag")
		End If
	End If

	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes,tableID)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		Else
			'If this is an insert, insert a record in the corresponding technical
			'data table.
			If Request("RECORD") < 0 Then
				tableName = LCase(GetEquipmentTypeName(Request("equipment_type_id"))) & "_technical_data"
				idName = "technical_data_id"
				ReDim fieldNames(0)
				ReDim fieldVals(UBound(fieldNames))
				ReDim fieldTypes(UBound(fieldNames))
				fieldNames(0) = "equipment_item_id"
				For count = 0 To UBound(fieldNames)
					fieldVals(count) = tableID
				Next
				fieldTypes(0) = "Number"
				status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes,tableID)
				If Not status Then
					session("err") = "db"
					For count = 0 To UBound(fieldNames)
						Session(fieldNames(count)) = fieldVals(count)
					Next
				End If
			End If
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"equipmenttypes") > 0 Then
	tableName = "equipment_types"
	idName = "equipment_type_id"
	ReDim fieldNames(3)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "equipment_type_name"
	fieldNames(1) = "equipment_type_description"
	fieldNames(2) = "inspection_interval"
	fieldNames(3) = "inspection_interval_units"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Text"
	fieldTypes(1) = "Text"
	fieldTypes(2) = "Number"
	fieldTypes(3) = "Text"

	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("equipment_type_name") = "" Then
			session("err") = "equipment_type_name"
			session("saveval") = request("equipment_type_name")
		ElseIf request("equipment_type_description") = "" Then
			session("err") = "equipment_type_description"
			session("saveval") = request("equipment_type_description")
		End If
	End If

	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes,tableID)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

End If

'Pop up a message and return to the calling page.
If session("err") = "db" Then
'	Response.Write "<script language='javascript'>alert('A database error occurred during update: " & FixString(Err.Description) & "');</script>"
	Response.Write "<script language='javascript'>alert('A database error occurred during update: " & FixString(Err.Description) & "'); window.location.href='" & Request("http_referer") & "';</script>"
ElseIf session("err") <> "" And session("err") <> "NONE" Then
	Response.Write "<script language='javascript'>alert('One or more errors occurred during update. Field [" & session("err") & "] is a required field or has an invalid value.'); window.location.href='" & Request("http_referer") & "';</script>"
Else
	If samePage = True Then
		Response.Write "<script language='javascript'>window.location.href='" & Left(Request("http_referer"),InStr(Request("http_referer"),"?") - 1) & "?record_id=" & NewID & "';</script>"
	ElseIf InStr(Request("http_referer"),"?") > 0 Then
		If InStr(Request("http_referer"),"sort=") > 0 Then
			Response.Write "<script language='javascript'>window.location.href='" & Left(Request("http_referer"),InStr(Request("http_referer"),"?") - 1) & "?sort=" & Request("SORT") & "&direction=" & Request("DIRECTION") & "&limit=" & Request("limit") & "';</script>"
		Else
			Response.Write "<script language='javascript'>window.location.href='" & Left(Request("http_referer"),InStr(Request("http_referer"),"?") - 1) & "';</script>"
		End If
	Else
		Response.Write "<script language='javascript'>window.location.href='" & Request("http_referer") & "';</script>"
	End If
End If

%>

<body bgcolor='#d0d0d0' link='black' vLink='black'>
</body>
</html>