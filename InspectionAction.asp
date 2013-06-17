<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<html>
<head>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Inspection Action</title>
<!--#include file="InspectionsFunctions.asp"-->
</head>
<%
Function WriteAuditInsertQuery(user,idVal,tableName,fieldName,oldVal,newVal,auditType)
	WriteAuditInsertQuery = "INSERT INTO audit_trail (change_user,change_item_id,change_table,change_field,old_value,new_value,change_type) " & _
					"VALUES ('" & user & "','" & idVal & "','" & tableName & "','" & fieldName & "','" & oldVal & "','" & newVal & "','" & auditType & "')"
End Function

'*************
' Revision History
' 
' Keith Brooks - Thursday, January 27, 2011
'   Creation.
'*************

Dim sqlString
Dim sqlString2
Dim cn
Dim cnInfo
Dim rs
Dim rsInfo
Dim equipType
Dim fieldName
Dim fieldType
Dim updateString
Dim valueString
Dim fieldString
Dim item_id
Dim inspection_id
Dim errFlag
Dim oldValue
Dim newValue
Dim currentuser
Dim auditFlag

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

errFlag = False
Session("focus") = ""

If Request("equipType") <> "" Then
	equipType = Request("equipType")
	
	set cn = CreateObject("adodb.connection")
	cn.Open = DBString
	Set cnInfo = CreateObject("adodb.connection")
	cnInfo.Open = InfoDBString
	Set rs = CreateObject("adodb.recordset")
	Set rsInfo = CreateObject("adodb.recordset")

	'Get the fields and datatypes for the appropriate inspection data table from
	'the information_schema database.
	sqlString = "SELECT column_name,data_type FROM columns " & _
			"WHERE table_schema='inspections' AND table_name='" & _
			equipType & "_inspection_data'"
	Set rsInfo = cnInfo.Execute(sqlString)
	If Not rsInfo.BOF Then
		cn.BeginTrans
		rsInfo.MoveFirst
		updateString = ""
		valueString = ""
		sqlString2 = ""
		
		'If the inspectionID request object exists and is > 0, then this is an update
		'to an existing inspection record; otherwise, insert the record.
		If CLng(Request("inspectionID")) > 0 Then
			inspection_id = Request("inspectionID")
			'Get the current record values for the audit trail.
			sqlString2 = "SELECT * FROM " & equipType & "_inspection_data " & _
					"WHERE inspection_data_id=" & Request("inspectionID")
			Set rs = cn.Execute(sqlString2)
			If Not rs.BOF Then
				rs.MoveFirst
				Do While Not rsInfo.EOF
					fieldName = rsInfo("column_name")
					fieldType = rsInfo("data_type")
					
					'Disregard fields ending in "id" because they are not updateable.
					If Right(fieldName,2) <> "id" Then
			
						'Save the current record value for the audit trail.
						If Not IsNull(rs(fieldName)) Then
							oldValue = CStr(rs(fieldName))
						Else
							oldValue = ""
						End If

						'Generate the SQL to update the existing inspection record.
						fieldString = ""
						If Request(fieldName) = "" Or ((fieldType = "datetime" Or fieldType = "date") And UCase(Request(fieldName)) = "NONE") Then
							If fieldType = "tinyint" Then
								fieldString = fieldName & "=false"
								newValue = "0"
							Else
								fieldString = fieldName & "=null"
								newValue = ""
							End If
						Else
							Select Case fieldType
								Case "datetime"
									fieldString = fieldName & "='" & FormatMySQLDateTime(Request(fieldName)) & "'"
									newValue = Request(fieldName)
								Case "date"
									fieldString = fieldName & "='" & FormatMySQLDate(Request(fieldName)) & "'"
									newValue = Request(fieldName)
								Case "int","smallint","double","float"
									fieldString = fieldName & "=" & Request(fieldName)
									newValue = CStr(Request(fieldName))
								Case "varchar","text"
									fieldString = fieldName & "='" & FixString(Request(fieldName)) & "'"
									newValue = Request(fieldName)
								Case "tinyint"
									fieldString = fieldName & "=true"
									newValue = "1"
							End Select
						End If
						If fieldString <> "" Then
							If updateString = "" Then
								updateString = fieldString
							Else
								updateString = updateString & "," & fieldString
							End If
						End If
								
						'If the value has changed, insert a record into the audit trail.
						If CStr(oldValue) <> CStr(newValue) Then
							cn.Execute(WriteAuditInsertQuery(currentuser,rs("equipment_item_id"),equipType & "_inspection_data",fieldName,oldValue,FixString(newValue),"update"))
							If Err.number <> 0 Then
								Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
								errFlag = True
							End If
						End If
					End If
					rsInfo.MoveNext
				Loop
			End If
			rs.Close
			
			If updateString <> "" Then
				sqlString2 = "UPDATE " & equipType & "_inspection_data SET " & _
						updateString & " WHERE inspection_data_id=" & Request("inspectionID")
			End If
			'Execute the SQL.
			If sqlString2 <> "" Then
				Err.Clear
				On Error Resume Next
				cn.Execute(sqlString2)
				If Err.number <> 0 Then
					Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
					errFlag = True
				End If
			Else
				errFlag = True
			End If
		
		'If the inspectionID doesn't exist, insert the record.
		ElseIf CLng(Request("itemID")) > 0 Then
			Do While Not rsInfo.EOF
				fieldName = rsInfo("column_name")
				fieldType = rsInfo("data_type")
				
				'Disregard fields ending in "id" because they are not updateable.
				If Right(fieldName,2) <> "id" Then
			
					'Generate the SQL to insert a new inspection record.
					fieldString = ""
					If Request(fieldName) = "" Or ((fieldType = "datetime" Or fieldType = "date") And UCase(Request(fieldName)) = "NONE") Then
						If fieldType = "tinyint" Then
							fieldString = "false"
						Else
							fieldString = "null"
						End If
						auditFlag = False
					Else
						Select Case fieldType
							Case "datetime"
								fieldString = "'" & FormatMySQLDateTime(Request(fieldName)) & "'"
							Case "date"
								fieldString = "'" & FormatMySQLDate(Request(fieldName)) & "'"
							Case "int","smallint","double","float"
								fieldString = Request(fieldName)
							Case "varchar","text"
								fieldString = "'" & FixString(Request(fieldName)) & "'"
							Case "tinyint"
								fieldString = "true"
						End Select
						auditFlag = True
					End If
					If fieldString <> "" Then
						updateString = updateString & "," & fieldName
						valueString = valueString & "," & fieldString
					End If
							
					'Insert a record into the audit trail if a value was specified.
					If auditFlag = True Then
						cn.Execute(WriteAuditInsertQuery(currentuser,Request("itemID"),equipType & "_inspection_data",fieldName,"null",FixString(Request(fieldName)),"insert"))
						If Err.number <> 0 Then
							Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
							errFlag = True
						End If
					End If
							
				End If
				rsInfo.MoveNext
			Loop
			
			sqlString2 = ""
			If updateString <> "" Then
				sqlString2 = "INSERT INTO " & equipType & "_inspection_data " & _
						"(equipment_item_id" & updateString & ") VALUES (" & _
						Request("itemID") & valueString & ")"
			End If
			'Execute the SQL.
			If sqlString2 <> "" Then
				Err.Clear
				On Error Resume Next
				cn.Execute(sqlString2)
				If Err.number <> 0 Then
					Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
					errFlag = True
				End If
			Else
				errFlag = True
			End If
			'Get the id number assigned to the new inspection.
			sqlString2 = "SELECT LAST_INSERT_ID()"
			Set rs = cn.Execute(sqlString2)
			If Not rs.BOF Then
				rs.MoveFirst
				inspection_id = rs(0)
			Else
				inspection_id = 0
			End If
			rs.Close
		End If
		
		'Update the technical data with the next inspection information.
		'Get the equipment_item_id.
		If errFlag = False Then
			item_id = 0
			If CLng(Request("itemID")) > 0 Then
				item_id = Request("itemID")
			ElseIf CLng(Request("inspectionID")) > 0 Then
				sqlString2 = "SELECT equipment_item_id FROM " & _
						equipType & "_inspection_data " & _
						"WHERE inspection_data_id=" & Request("inspectionID")
				Set rs = cn.Execute(sqlString2)
				If Not rs.BOF Then
					rs.MoveFirst
					item_id = rs(0)
				End If
				rs.Close
			End If
		
			If item_id > 0 Then
				'Make sure that the fields exist and build the update string.
				updateString = ""
				If Request("set_frequency") <> "" Then
					updateString = "inspection_frequency=" & Request("set_frequency")
				End If
				If Request("set_frequency_units") <> "" Then
					If updateString = "" Then
						updateString = "inspection_frequency_units='" & Request("set_frequency_units") & "'"
					Else
						updateString = updateString & ",inspection_frequency_units='" & Request("set_frequency_units") & "'"
					End If
				End If
				If Request("next_inspection_due") <> "" Then
					If updateString = "" Then
						updateString = "next_inspection_date='" & FormatMySQLDate(Request("next_inspection_due")) & "'"
					Else
						updateString = updateString & ",next_inspection_date='" & FormatMySQLDate(Request("next_inspection_due")) & "'"
					End If
				End If
				If updateString <> "" Then
					sqlString2 = "UPDATE " & equipType & "_technical_data SET " & _
								updateString & " WHERE equipment_item_id=" & item_id
					Err.Clear
					On Error Resume Next
					cn.Execute(sqlString2)
					If Err.number <> 0 Then
						Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
						errFlag = True
					End If
				End If
			End If
		End If
		
		'If action = "updateReading" and the reading id > 0, update the
		'appropriate inspection_readings record.
		If Request("action") = "updateReading" And Request("readingID") <> "" Then
			If CLng(Request("readingID")) > 0 Then
				sqlString2 = "UPDATE " & equipType & "_inspection_readings SET " & _
						"location='" & Request("edit_location") & "'," & _
						"reading=" & Request("edit_reading") & _
						" WHERE inspection_reading_id=" & Request("readingID")
			'If the reading id <= 0 Then insert the record.
			Else
				sqlString2 = "INSERT INTO " & equipType & "_inspection_readings " & _
						"(inspection_data_id,location,reading) VALUES (" & _
						inspection_id & ",'" & Request("edit_location") & _
						"'," & Request("edit_reading") & ")"
			End If
			Err.Clear
			On Error Resume Next
			cn.Execute(sqlString2)
			If Err.Number <> 0 Then
				Response.Write "<h2>DatabaseError - " & Err.Description & "</h2>"
				errFlag = True
			End If
		End If
			
		
		'If everything was successful, commit the transaction; otherwise,
		'rollback the transaction.
		If errFlag = False Then
			cn.CommitTrans
		Else
			cn.RollbackTrans
		End If

	Else
		Response.Write "<h2>Error - information table not found.</h2>"
		errFlag = True
	End If

	rsInfo.Close
	Set rsInfo = Nothing
	Set rs = Nothing
	cn.Close
	cnInfo.Close
	Set cn = Nothing
	Set cnInfo = Nothing
	
	'If everything was successful, go back to the appropriate page.
	If errFlag = False Then
		'If action = "editReading", go back to the inspection form with the
		'readingID specified.
		If Request("action") = "editReading" Then
			Session("focus") = "edit_location"
			Response.Write "<script language='javascript'>window.location.href='" & Left(Request("http_referer"),InStr(Request("http_referer"),"?") - 1) & "?inspectionID=" & inspection_id & "&readingID=" & Request("readingID") & "&edit=true';</script>"
		'If action = "editReading", go back to the inspection form with the
		'readingID specified as -1.
'		ElseIf Request("action") = "addReading" Then
'			Session("focus") = "edit_location"
'			Response.Write "<script language='javascript'>window.location.href='" & Left(Request("http_referer"),InStr(Request("http_referer"),"?") - 1) & "?inspectionID=" & inspection_id & "&readingID=-1&edit=true';</script>"
		'If action = "updateReading", go back to the inspection form.
		ElseIf Request("action") = "updateReading" Then
			Session("focus") = "edit_location"
			Response.Write "<script language='javascript'>window.location.href='" & Left(Request("http_referer"),InStr(Request("http_referer"),"?") - 1) & "?inspectionID=" & inspection_id & "&edit=true';</script>"
		'Otherwise, go back to the main manu.
		Else
			Response.Redirect("default.asp")
		End If
	End If

Else
	Response.Write "<h2>Error - Equipment type not specified</h2>"
End If
%>

<body>
</body>
</html>