<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<html>
<head>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Technical Data Action</title>
<!--#include file="InspectionsFunctions.asp"-->
</head>
<%
Function WriteAuditInsertQuery(user,idVal,tableName,fieldName,oldVal,newVal,auditType)
	WriteAuditInsertQuery = "INSERT INTO admin_audit_trail (change_user,change_table_id,change_table,change_field,old_value,new_value,change_type) " & _
					"VALUES ('" & user & "','" & idVal & "','" & tableName & "','" & fieldName & "','" & oldVal & "','" & newVal & "','" & auditType & "')"
End Function

'*************
' Revision History
' 
' Keith Brooks - Thursday, February 10, 2011
'   Creation.
'*************

Dim sqlString
Dim sqlString2
Dim cn
Dim cnInfo
Dim rs
Dim rsInfo
Dim equipType
Dim equipTypeID
Dim plantarea
Dim fieldName
Dim fieldType
Dim updateString
Dim valueString
Dim fieldString
Dim errFlag
Dim oldValue
Dim newValue
Dim currentuser
Dim cv

errFlag = False

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'If the previous page was the "selectitem" form, determine the equipment type
'for the specified item ID, then forward to the appropriate technical data form.
If InStr(LCase(request("http_referer")),"selectitem") > 0 Then
	If Request("itemID") <> "" Then
		Set cn = CreateObject("adodb.connection")
		cn.Open = DBString
		Set rs = CreateObject("adodb.recordset")

		sqlString = "SELECT equipment_type_name,conservation_vent, " & _
				"a.equipment_type_id " & _
				"FROM equipment_types a " & _
				"INNER JOIN equipment_items b " & _
				"ON a.equipment_type_id=b.equipment_type_id " & _
				"WHERE b.equipment_item_id=" & Request("itemID")
		Set rs = cn.Execute(sqlString)
		equipType = ""
		equipTypeID = 0
		If Not rs.BOF Then
			rs.MoveFirst
			If Not IsNull(rs(0)) And rs(0) <> "" Then
				equipType = rs(0)
				equipTypeID = rs(2)
			End If
			If Not IsNull(rs(1)) Then
				cv = rs(1)
			Else
				cv = 0
			End If
		End If
		rs.Close

		Set rs = Nothing
		cn.Close
		Set cn = Nothing

		If equipType <> "" Then
			If LCase(equipType) = "psv" And cv = 1 Then
				Response.Redirect("cv_technicaldata.asp?itemID=" & Request("itemID") & "&edit=true&plantarea=" & Request("plantarea") & "&equipTypeID=" & equipTypeID)
			Else
				Response.Redirect(LCase(equipType) & "_technicaldata.asp?itemID=" & Request("itemID") & "&edit=true&plantarea=" & Request("plantarea") & "&equipTypeID=" & equipTypeID)
			End If
		Else
			Response.Write "<h2>Error - Unable to determine equipment type</h2>"
		End If
	Else
		Response.Write "<h2>Error - Equipment Item ID not specified</h2>"
	End If
Else
	'If the technical_data_id request object exists and is > 0, then this is an update
	'to an existing technical data record; otherwise, insert the record.
	If Request("equipType") <> "" Then
		equipType = Request("equipType")
		plantarea = Request("plantarea")
		equipTypeID = Request("equipTypeID")
		
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
				equipType & "_technical_data'"
		Set rsInfo = cnInfo.Execute(sqlString)
		If Not rsInfo.BOF Then
			cn.BeginTrans
			rsInfo.MoveFirst
			updateString = ""
			valueString = ""
			sqlString2 = ""
			If CLng(Request("technical_data_id")) > 0 Then
				'Get the current record values for the audit trail.
				sqlString2 = "SELECT * FROM " & equipType & "_technical_data WHERE technical_data_id=" & Request("technical_data_id")
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
							If Request(fieldName) = "" Then
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
								cn.Execute(WriteAuditInsertQuery(currentuser,Request("technical_data_id"),equipType & "_technical_data",fieldName,oldValue,newValue,"update"))
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
					sqlString2 = "UPDATE " & equipType & "_technical_data SET " & _
							updateString & " WHERE technical_data_id=" & Request("technical_data_id")
				Else
					sqlString2 = ""
				End If
			ElseIf CLng(Request("itemID")) > 0 Then
				Do While Not rsInfo.EOF
					fieldName = rsInfo("column_name")
					fieldType = rsInfo("data_type")
					
					'Disregard fields ending in "id" because they are not updateable.
					If Right(fieldName,2) <> "id" Then
				
						'Generate the SQL to update the existing inspection record.
						fieldString = ""
						If Request(fieldName) = "" Then
							If fieldType = "tinyint" Then
								fieldString = "false"
							Else
								fieldString = "null"
							End If
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
						End If
						If fieldString <> "" Then
							updateString = updateString & "," & fieldName
							valueString = valueString & "," & fieldString
						End If
						
						'Insert a record into the audit trail.
						cn.Execute(WriteAuditInsertQuery(currentuser,0,equipType & "_technical_data",fieldName,"null",Request(fieldName),"insert"))
						If Err.number <> 0 Then
							Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
							errFlag = True
						End If
								
					End If
					rsInfo.MoveNext
				Loop
				
				sqlString2 = ""
				If updateString <> "" Then
					sqlString2 = "INSERT INTO " & equipType & "_technical_data " & _
							"(equipment_item_id" & updateString & ") VALUES (" & _
							Request("itemID") & valueString & ")"
				End If
			
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

			'If everything was successful, commit the transaction; otherwise,
			'rollback the transaction.
			If errFlag = False Then
				cn.CommitTrans
			Else
				cn.RollbackTrans
			End If

		Else
			Response.Write "<h2>Error - information table not found.</h2>"
		End If

		rsInfo.Close
		Set rsInfo = Nothing
		Set rs = Nothing
		cn.Close
		cnInfo.Close
		Set cn = Nothing
		Set cnInfo = Nothing

		'If everything was successful, go back to the select item page.
		If errFlag = False Then
			Response.Redirect("selectitem.asp?plantarea=" & plantarea & "&equipType=" & equipTypeID & "&form_action=edittechnicaldata&flowflag=true")
		End If

	Else
		Response.Write "<h2>Error - Equipment type not specified</h2>"
	End If
End If
%>

<body>
</body>
</html>