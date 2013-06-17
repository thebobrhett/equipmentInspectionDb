<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<html>
<head>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>PM Action</title>
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
' Keith Brooks - Friday, January 6, 2012
'   Creation.
'*************

Dim sqlString
Dim cn
Dim rs
Dim rs2
Dim errVar
Dim errRecord
Dim focusVar
Dim pm_data_id
Dim sql
Dim errFlag
Dim updateString
Dim valueString
Dim oldValue
Dim newValue
Dim temp
Dim pm_dateField

Dim currentuser
'Dim auditFlag
Dim tableID
Dim tableID2

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Make sure that an equipment_item_id has been specified.
If IsNumeric(Request("equipment_item_id")) Then
	errFlag = False
	focusVar = ""
	errVar = ""
	errRecord = ""
	'Calculate the name of the pm_date field (done to allow the datepicker
	'on the main form to work.
	pm_dateField = "pm_date" & Request("equipment_item_id")

	'If this is a submit after an error and this is for the same record, specify
	'the element to receive the focus when we return.
	'If IsNumeric(errRecord) And errVar <> "" Then
	'	If CLng(errRecord) = CLng(Request("equipment_item_id")) Then
	'		If errVar = pm_dateField Then
	'			focus = 

	'Save the variable values in session variables to maintain the entered data
	'if there is an error.
	Session(pm_dateField) = Request(pm_dateField)
	Session("in_need_of_repair") = Request("in_need_of_repair")
	Session("seal_condition") = Request("seal_condition")
	Session("comments") = Request("comments")
	Session("front_bearing_db_level") = Request("front_bearing_db_level")
	Session("rear_bearing_db_level") = Request("rear_bearing_db_level")
	Session("top_db_level") = Request("top_db_level")
	Session("spare1") = Request("spare1")
	Session("spare2") = Request("spare2")
	Session("pm_data_id") = Request("pm_data_id")
	Session("pm_subitem1_data_id") = Request("pm_subitem1_data_id")
	Session("pm_subitem2_data_id") = Request("pm_subitem2_data_id")

	'Check for required fields and data validity.
	If Request("top_db_level") <> "" Then
		If Not IsNumeric(Request("top_db_level")) Then
			errVar = "top_db_level"
			errRecord = Request("equipment_item_id")
		End If
	End If
	If Request("rear_bearing_db_level") <> "" Then
		If Not IsNumeric(Request("rear_bearing_db_level")) Then
			errVar = "rear_bearing_db_level"
			errRecord = Request("equipment_item_id")
		End If
	End If
	If Request("front_bearing_db_level") <> "" Then
		If Not IsNumeric(Request("front_bearing_db_level")) Then
			errVar = "front_bearing_db_level"
			errRecord = Request("equipment_item_id")
		End If
	End If
	If Not IsDate(Request(pm_dateField)) Then
		errVar = pm_dateField
		errRecord = Request("equipment_item_id")
	End If
	Session("errVar") = errVar
	Session("errRecord") = errRecord

	'If no error has occurred, insert/update the records and write the changes
	'to the audit trail.
	If errVar = "" Then
	
		'Create the database objects.	
		Set cn = CreateObject("adodb.connection")
		cn.Open = DBString
		cn.BeginTrans
		Set rs = CreateObject("adodb.recordset")
		Set rs2 = CreateObject("adodb.recordset")

		'If the pm_data_id is set to -1, insert new records for the data;
		'otherwise, update the existing reccords.
		If IsNumeric(Request("pm_data_id")) Then
			pm_data_id = CLng(Request("pm_data_id"))
		Else
			pm_data_id = 0
		End If
		If pm_data_id > 0 Then
			'Get the existing values from the pm_data table for the audit trail.
			sql = "SELECT * FROM pm_data WHERE pm_data_id=" & pm_data_id
			Set rs = cn.Execute(sql)
			If Not rs.BOF Then
				rs.MoveFirst
				updateString = ""
				'If a field has changed, insert a record into the audit trail.
				If CDate(rs("pm_date")) <> CDate(Request(pm_dateField)) Then
					updateString = "pm_date='" & FormatMySQLDate(Request(pm_dateField)) & "'"
					cn.Execute(WriteAuditInsertQuery(currentuser,pm_data_id,"pm_data","pm_date",rs("pm_date"),FixString(Request(pm_dateField)),"update"))
					If Err.number <> 0 Then
						Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
						errFlag = True
					End If
				End If
				If CBool(rs("in_need_of_repair")) <> CBool(Request("in_need_of_repair")) Then
					If CBool(Request("in_need_of_repair")) Then
						If updateString = "" Then
							updateString = "in_need_of_repair=1"
						Else
							updateString = updateString & ",in_need_of_repair=1"
						End If
					Else
						If updateString = "" Then
							updateString = "in_need_of_repair=0"
						Else
							updateString = updateString & ",in_need_of_repair=0"
						End If
					End If
					cn.Execute(WriteAuditInsertQuery(currentuser,pm_data_id,"pm_data","in_need_of_repair",rs("in_need_of_repair"),FixString(Request("in_need_of_repair")),"update"))
					If Err.number <> 0 Then
						Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
						errFlag = True
					End If
				End If
				If rs("seal_condition") <> FixString(Request("seal_condition")) Then
					If updateString = "" Then
						updateString = "seal_condition='" & Request("seal_condition") & "'"
					Else
						updateString = updateString & ",seal_condition='" & Request("seal_condition") & "'"
					End If
					cn.Execute(WriteAuditInsertQuery(currentuser,pm_data_id,"pm_data","seal_condition",FixString(rs("seal_condition")),FixString(Request("seal_condition")),"update"))
					If Err.number <> 0 Then
						Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
						errFlag = True
					End If
				End If
				If rs("comments") <> FixString(Request("comments")) Then
					If updateString = "" Then
						updateString = "comments='" & FixString(Request("comments")) & "'"
					Else
						updateString = updateString & ",comments='" & FixString(Request("comments")) & "'"
					End If
					cn.Execute(WriteAuditInsertQuery(currentuser,pm_data_id,"pm_data","comments",FixString(rs("comments")),FixString(Request("comments")),"update"))
					If Err.number <> 0 Then
						Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
						errFlag = True
					End If
				End If
				rs.Close
				
				If Not errFlag And UpdateString <> "" Then
					Err.Clear
					'Write the update for the pm_data table.
					sql = "UPDATE pm_data SET " & updateString & _
						" WHERE pm_data_id=" & pm_data_id
					cn.Execute(sql)
					If Err.number <> 0 Then
						Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
						errFlag = True
					End If
				End If
				
				'Get the existing values from the pm_subitem_data table for the audit trail.
				If IsNumeric(Request("pm_subitem1_data_id")) Then
					sql = "SELECT * FROM pm_subitem_data WHERE pm_subitem_data_id=" & Request("pm_subitem1_data_id")
					Set rs = cn.Execute(sql)
					If Not rs.BOF Then
						rs.MoveFirst
						updateString = ""
						'If a field has changed, insert a record into the audit trail.
						If CStr(Nz(rs("front_bearing_db_level"),"")) <> CStr(Nz(Request("front_bearing_db_level"),"")) Then
							If Request("front_bearing_db_level") = "" Then
								updateString = "front_bearing_db_level=null"
								cn.Execute(WriteAuditInsertQuery(currentuser,Request("pm_subitem1_data_id"),"pm_subitem_data","front_bearing_db_level",rs("front_bearing_db_level"),"null","update"))
							Else
								updateString = "front_bearing_db_level=" & Request("front_bearing_db_level")
								cn.Execute(WriteAuditInsertQuery(currentuser,Request("pm_subitem1_data_id"),"pm_subitem_data","front_bearing_db_level",rs("front_bearing_db_level"),Request("front_bearing_db_level"),"update"))
							End If
							If Err.number <> 0 Then
								Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
								errFlag = True
							End If
						End If
						If CStr(Nz(rs("rear_bearing_db_level"),"")) <> CStr(Nz(Request("rear_bearing_db_level"),"")) Then
							If Request("rear_bearing_db_level") = "" Then
								If updateString = "" Then
									updateString = "rear_bearing_db_level=null"
								Else
									updateString = updateString & ",rear_bearing_db_level=null"
								End If
								cn.Execute(WriteAuditInsertQuery(currentuser,Request("pm_subitem1_data_id"),"pm_subitem_data","rear_bearing_db_level",rs("rear_bearing_db_level"),"null","update"))
							Else
								If updateString = "" Then
									updateString = "rear_bearing_db_level=" & Request("rear_bearing_db_level")
								Else
									updateString = updateString & ",rear_bearing_db_level=" & Request("rear_bearing_db_level")
								End If
								cn.Execute(WriteAuditInsertQuery(currentuser,Request("pm_subitem1_data_id"),"pm_subitem_data","rear_bearing_db_level",rs("rear_bearing_db_level"),Request("rear_bearing_db_level"),"update"))
							End If
							If Err.number <> 0 Then
								Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
								errFlag = True
							End If
						End If
						If CStr(Nz(rs("spare"),"")) <> CStr(Nz(Request("spare1"),"")) Then
							If Request("spare1") = "1" Then
								If updateString = "" Then
									updateString = "spare=1"
								Else
									updateString = updateString & ",spare=1"
								End If
							Else
								If updateString = "" Then
									updateString = "spare=0"
								Else
									updateString = updateString & ",spare=0"
								End If
							End If
							cn.Execute(WriteAuditInsertQuery(currentuser,Request("pm_subitem1_data_id"),"pm_subitem_data","spare",rs("spare"),FixString(Request("spare1")),"update"))
							If Err.number <> 0 Then
								Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
								errFlag = True
							End If
						End If
						rs.Close
								
						If Not errFlag And updateString <> "" Then
							Err.Clear
							'Write the first update string for the
							'pm_subitem_data table.
							sql = "UPDATE pm_subitem_data SET " & updateString & _
								" WHERE pm_subitem_data_id=" & Request("pm_subitem1_data_id")
							'Response.Write "sql = " & sql
							cn.Execute(sql)
							If Err.number <> 0 Then
								Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
								errFlag = True
							End If
						End If
					Else
						Response.Write "<h2>Database Error - Existing record not found!</h2>"
						errFlag = True
						rs.Close
					End If
				End If
				'Get the existing values for the 2nd item from the pm_subitem_data table for the audit trail.
				If IsNumeric(Request("pm_subitem2_data_id")) Then
					sql = "SELECT * FROM pm_subitem_data WHERE pm_subitem_data_id=" & Request("pm_subitem2_data_id")
					Set rs = cn.Execute(sql)
					If Not rs.BOF Then
						rs.MoveFirst
						updateString = ""
						'If a field has changed, insert a record into the audit trail.
						If CStr(Nz(rs("top_db_level"),"")) <> CStr(Nz(Request("top_db_level"),"")) Then
							If Request("top_db_level") = "" Then
								updateString = "top_db_level=null"
								cn.Execute(WriteAuditInsertQuery(currentuser,Request("pm_subitem2_data_id"),"pm_subitem_data","top_db_level",rs("top_db_level"),"null","update"))
							Else
								updateString = "top_db_level=" & Request("top_db_level")
								cn.Execute(WriteAuditInsertQuery(currentuser,Request("pm_subitem2_data_id"),"pm_subitem_data","top_db_level",rs("top_db_level"),Request("top_db_level"),"update"))
							End If
							If Err.number <> 0 Then
								Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
								errFlag = True
							End If
						End If
						If Request("spare2") <> "" Then
							If CStr(Nz(rs("spare"),"")) <> CStr(Nz(Request("spare2"),"")) Then
								If CStr(Nz(Request("spare2"),"")) = "1" Then
									If updateString = "" Then
										updateString = "spare=1"
									Else
										updateString = updateString & ",spare=1"
									End If
								Else
									If updateString = "" Then
										updateString = "spare=0"
									Else
										updateString = updateString & ",spare=0"
									End If
								End If
								cn.Execute(WriteAuditInsertQuery(currentuser,Request("pm_subitem2_data_id"),"pm_subitem_data","spare",rs("spare"),Request("spare2"),"update"))
								If Err.number <> 0 Then
									Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
									errFlag = True
								End If
							End If
						Else
							If CBool(rs("spare")) = True Then
								If updateString = "" Then
									updateString = "spare=0"
								Else
									updateString = updateString & ",spare=0"
								End If
								cn.Execute(WriteAuditInsertQuery(currentuser,Request("pm_subitem2_data_id"),"pm_subitem_data","spare",rs("spare"),"0","update"))
								If Err.number <> 0 Then
									Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
									errFlag = True
								End If
							End If
						End If
						rs.Close
								
						If Not errFlag And updateString <> "" Then
							Err.Clear
							'Write the second update string for the
							'pm_subitem_data table.
							sql = "UPDATE pm_subitem_data SET " & updateString & _
								" WHERE pm_subitem_data_id=" & Request("pm_subitem2_data_id")
							'Response.Write "sql = " & sql
							cn.Execute(sql)
							If Err.number <> 0 Then
								Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
								errFlag = True
							End If
						End If
					Else
						Response.Write "<h2>Database Error - Existing record not found!</h2>"
						errFlag = True
						rs.Close
					End If
				End If
			Else
				Response.Write "<h2>Database Error - Existing record not found!</h2>"
				errFlag = True
				rs.Close
			End If
		Else
			'Write the query to insert the inspection into the pm_data table.
			valueString = Request("equipment_item_id")
			If Not IsDate(Request(pm_dateField)) Then
				valueString = valueString & ",'" & FormatMySQLDate(Request(pm_dateField)) & "'"
			Else
				valueString = valueString & ",'" & FormatMySQLDate(Date) & "'"
			End If
			If CBool(Request("in_need_of_repair")) = True Then
				valueString = valueString & ",1"
			Else
				valueString = valueString & ",0"
			End If
			If Request("seal_condition") <> "" Then
				valueString = valueString & ",'" & Request("seal_condition") & "'"
			Else
				valueString = valueString & ",null"
			End If
			If Request("comments") <> "" Then
				valueString = valueString & ",'" & Request("comments") & "'"
			Else
				valueString = valueString & ",null"
			End if
			
			sql = "INSERT INTO pm_data (equipment_item_id,pm_date," & _
				"in_need_of_repair,seal_condition,comments) VALUES " & _
				"(" & valueString & ")"
			cn.Execute(sql)
			If Err.number <> 0 Then
				Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
				errFlag = True
			Else
				'Get the index number that was just assigned.
				sql = "SELECT LAST_INSERT_ID()"
				Set rs = cn.Execute(sql)
				If Err.number <> 0 Then
					Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
					errFlag = True
				End If
				If Not rs.BOF Then
					rs.MoveFirst
					tableID = rs(0)
				Else
					tableID = 0
				End If
				rs.Close
				
				'Write the insert values to the audit trail.
				cn.Execute(WriteAuditInsertQuery(currentuser,tableID,"pm_data","pm_data_id","null",tableID,"insert"))
				cn.Execute(WriteAuditInsertQuery(currentuser,tableID,"pm_data","equipment_item_id","null",Request("equipment_item_id"),"insert"))
				cn.Execute(WriteAuditInsertQuery(currentuser,tableID,"pm_data","pm_date","null",FixString(Request(pm_dateField)),"insert"))
				cn.Execute(WriteAuditInsertQuery(currentuser,tableID,"pm_data","in_need_of_repair","null",Request("in_need_of_repair"),"insert"))
				cn.Execute(WriteAuditInsertQuery(currentuser,tableID,"pm_data","seal_condition","null",FixString(Request("seal_condition")),"insert"))
				cn.Execute(WriteAuditInsertQuery(currentuser,tableID,"pm_data","comments","null",FixString(Request("comments")),"insert"))
				
				'Write the query for the first record to be inserted into the
				'pm_subitem_data table.
				sql = "SELECT * FROM equipment_items_subitems " & _
					"WHERE equipment_item_id=" & Request("equipment_item_id") & _
					" ORDER BY equipment_items_subitem_id"
				Set rs = cn.Execute(sql)
				If Err.number <> 0 Then
					Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
					errFlag = True
				End If
				If Not rs.BOF Then
					rs.MoveFirst
					Do While Not rs.EOF
						valueString = tableID
						valueString = valueString & "," & rs("equipment_items_subitem_id")
						If CBool(rs("test_front_bearing")) = True Then
							valueString = valueString & "," & Request("front_bearing_db_level")
						Else
							valueString = valueString & ",null"
						End If
						If CBool(rs("test_rear_bearing")) = True Then
							valueString = valueString & "," & Request("rear_bearing_db_level")
						Else
							valueString = valueString & ",null"
						End If
						If CBool(rs("test_top")) = True Then
							valueString = valueString & "," & Request("top_db_level")
						Else
							valueString = valueString & ",null"
						End If
						If CBool(rs("test_front_bearing")) = True Then
							If Request("spare1") <> "" Then
								valueString = valueString & "," & Request("spare1")
							Else
								valueString = valueString & ",0	"
							End If
						Else
							If Request("spare2") <> "" Then
								valueString = valueString & "," & Request("spare2")
							Else
								valueString = valueString & ",0"
							End If
						End If
						sql = "INSERT INTO pm_subitem_data (pm_data_id," & _
							"equipment_items_subitem_id,front_bearing_db_level," & _
							"rear_bearing_db_level,top_db_level,spare) VALUES " & _
							"(" & valueString & ")"
						cn.Execute(sql)
						If Err.number <> 0 Then
							Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
							errFlag = True
						End If
						
						'Get the index number that was just assigned.
						sql = "SELECT LAST_INSERT_ID()"
						Set rs2 = cn.Execute(sql)
						If Err.number <> 0 Then
							Response.Write "<h2>Database Error - " & Err.Description & "</h2>"
							errFlag = True
						End If
						If Not rs2.BOF Then
							rs2.MoveFirst
							tableID2 = rs(0)
						Else
							tableID2 = 0
						End If
						rs2.Close
						
						'Write the insert values to the audit trail.
						cn.Execute(WriteAuditInsertQuery(currentuser,tableID2,"pm_subitem_data","pm_subitem_data_id","null",tableID2,"insert"))
						cn.Execute(WriteAuditInsertQuery(currentuser,tableID2,"pm_subitem_data","pm_data_id","null",tableID,"insert"))
						cn.Execute(WriteAuditInsertQuery(currentuser,tableID2,"pm_subitem_data","pm_subitem_data_id","null",tableID2,"insert"))
						If CBool(rs("test_front_bearing")) = True Then
							cn.Execute(WriteAuditInsertQuery(currentuser,tableID2,"pm_subitem_data","front_bearing_db_level","null",Request("front_bearing_db_level"),"insert"))
							cn.Execute(WriteAuditInsertQuery(currentuser,tableID2,"pm_subitem_data","spare","null",Request("spare1"),"insert"))
						End If
						If CBool(rs("test_rear_bearing")) = True Then
							cn.Execute(WriteAuditInsertQuery(currentuser,tableID2,"pm_subitem_data","rear_bearing_db_level","null",Request("rear_bearing_db_level"),"insert"))
						End If
						If CBool(rs("test_top")) = True Then
							cn.Execute(WriteAuditInsertQuery(currentuser,tableID2,"pm_subitem_data","top_db_level","null",Request("top_db_level"),"insert"))
							cn.Execute(WriteAuditInsertQuery(currentuser,tableID2,"pm_subitem_data","spare","null",Request("spare2"),"insert"))
						End If
							
						rs.MoveNext
					Loop
				End If
				rs.Close
			End If
		End If

		If errFlag Then
			cn.RollbackTrans
		Else
			cn.CommitTrans
		End If



		Set rs = Nothing
		Set rs2 = Nothing
		cn.Close
		Set cn = Nothing
	
	Else
		'Go back to the main page to display the error.
		Response.Redirect("PM_Inspection.asp")
	End If

	'If everything was successful, go back to the main page.
	If errFlag = False Then
		Response.Redirect("PM_Inspection.asp")
	Else
	End If

Else
	Response.Write "<h2>Error - Equipment Item ID not specified</h2>"
End If
%>

<body>
</body>
</html>