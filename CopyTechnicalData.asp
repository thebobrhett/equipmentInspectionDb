<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<% Server.ScriptTimeout=900 %>
<html>
<head>
<meta name="GENERATOR" Content="Microsoft Visual Studio 6.0">
<title>Copy Technical Data From Instrument Database</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="InspectionsFunctions.asp"-->
<link rel=STYLESHEET href='formstyle.css' type='text/css'>
</head>
<body>
<form id="form1" name="form1" action="CopyTechnicalData.asp" method="post">
<input type="hidden" id="copydata" name="copydata" value="true" />
<%
'*************
' Revision History
' 
' Keith Brooks - Thursday, January 13, 2011
'   Creation.
'*************

Dim sqlString
Dim sqlString2
Dim cnInspections
Dim cnInstruments
Dim cnInfo
Dim rs
Dim rs2
Dim rs3
Dim rs4
Dim criteria
Dim tagname
Dim equipnum
Dim currentuser
Dim access
Dim itemID

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("inspections", "copytechnicaldata", currentuser)
If access <> "none" Then

	If Request("copydata") <> "true" Then
		Response.Write "<h2>This will copy the technical data for the PSV elements in the Instrument database to the Inspection database.</h2>"
		Response.Write "<h3>Click the button below to proceed</h3>"
		Response.Write "<input type='submit' name='submit1' id='submit1' value='Copy Data' />"
	Else
		'Define the ado connections and recordset objects.
		Set cnInspections = CreateObject("adodb.connection")
		cnInspections.Open = DBString
		Set cnInstruments = CreateObject("adodb.connection")
		cnInstruments.Open = InstrDBString
		Set cnInfo = CreateObject("adodb.connection")
		cnInfo.Open = InfoDBString
		set rs = CreateObject("adodb.recordset")
		set rs2 = CreateObject("adodb.recordset")
		set rs3 = CreateObject("adodb.recordset")
		Set rs4 = CreateObject("adodb.recordset")

		Response.Write "<h3>Copying data...</h3>"
		
		'Get the PSV instruments from the instruments database.
		sqlString = "SELECT a.instr_id,REPLACE(a.instr_name,' ','') AS tag," & _
				"a.instr_desc,c.plant_area_num " & _
				"FROM (instruments a INNER JOIN instrument_function_types b " & _
				"ON a.instr_func_type_id=b.instr_func_type_id) " & _
				"INNER JOIN plant_areas c " & _
				"ON a.plant_area_id=c.plant_area_id " & _
				"WHERE b.instr_func_type_name='PSV' " & _
				"ORDER BY a.instr_name"
		Set rs = cnInstruments.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			Do While Not rs.EOF
				'Get the equipment item with a tag that matches this instrument.
				sqlString = "SELECT equipment_item_id FROM equipment_items " & _
						"WHERE equipment_item_tag='" & rs("tag") & "'"
				Set rs2 = cnInspections.Execute(sqlString)
				itemID = 0
				If Not rs2.BOF Then
					rs2.MoveFirst
					itemID = rs2("equipment_item_id")
				End If
				rs2.Close
				If itemID = 0 Then
					Response.Write "Inserting record for " & rs("tag") & "<br />"
				Else
					Response.Write "Updating record for " & rs("tag") & "<br />"
				End If
				
				'Get the values from the cross-reference table.
				sqlString = "SELECT * FROM equipment_instrument_xref " & _
						"WHERE technical_data_table='psv_technical_data'"
				Set rs2 = cnInspections.Execute(sqlString)
				If Not rs2.BOF Then
					Dim fieldList
					fieldList = ""
					Dim valueList
					valueList = ""
					rs2.MoveFirst
					Do While Not rs2.EOF
						'If the "instrument_sql" field exists, execute the
						'specified query; otherwise, look up the specified table
						'and field in the instruments database.
						If Not IsNull(rs2("instrument_sql")) Then
							'Replace the placeholder with the ID for this instrument.
							sqlString = Replace(rs2("instrument_sql"),"[instr_id]",rs("instr_id"))
						Else
							sqlString = "SELECT " & rs2("instrument_field") & _
									" FROM " & rs2("instrument_table") & _
									" WHERE instr_id=" & rs("instr_id")
						End If
						Set rs3 = cnInstruments.Execute(sqlString)
						If Not rs3.BOF Then
							rs3.MoveFirst
							'Get the datatype of the technical data field.
							sqlString = "SELECT data_type FROM columns " & _
									"WHERE table_schema='inspections' " & _
									"AND table_name='psv_technical_data' " & _
									"AND column_name='" & rs2("technical_data_field") & "'"
							Set rs4 = cnInfo.Execute(sqlString)
							If Not rs4.BOF Then
								rs4.MoveFirst
								If rs4("data_type") = "varchar" Then								
									'If the item doesn't exist in the inspections database,
									'create an INSERT string; otherwise, create an UPDATE string.
									If itemID = 0 Then
										If fieldList = "" Then
											fieldList = rs2("technical_data_field")
										Else
											fieldList = fieldList & "," & rs2("technical_data_field")
										End If
										
										If valueList = "" Then
											If IsNull(rs3(0)) Then
												valueList = "null"
											Else
												valueList = "'" & rs3(0) & "'"
											End If
										Else
											If IsNull(rs3(0)) Then
												valueList = valueList & ",null"
											Else
												valueList = valueList & ",'" & rs3(0) & "'"
											End If
										End If
									Else
										If fieldList = "" Then
											If IsNull(rs3(0)) Then
												fieldList = rs2("technical_data_field") & "=null"
											Else
												fieldList = rs2("technical_data_field") & "='" & rs3(0) & "'"
											End If
										Else
											If IsNull(rs3(0)) Then
												fieldList = fieldList & "," & rs2("technical_data_field") & "=null"
											Else
												fieldList = fieldList & "," & rs2("technical_data_field") & "='" & rs3(0) & "'"
											End If
										End If
									End If
								Else
									'If the item doesn't exist in the inspections database,
									'create an INSERT string; otherwise, create an UPDATE string.
									If itemID = 0 Then
										If fieldList = "" Then
											fieldList = rs2("technical_data_field")
										Else
											fieldList = fieldList & "," & rs2("technical_data_field")
										End If
											
										If valueList = "" Then
											If IsNull(rs3(0)) Then
												valueList = "null"
											Else
												valueList = rs3(0)
											End If
										Else
											If IsNull(rs3(0)) Then
												valueList = valueList & ",null"
											Else
												valueList = valueList & "," & rs3(0)
											End If
										End If
									Else
										If fieldList = "" Then
											If IsNull(rs3(0)) Then
												fieldList = rs2("technical_data_field") & "=null"
											Else
												fieldList = rs2("technical_data_field") & "=" & rs3(0)
											End If
										Else
											If IsNull(rs3(0)) Then
												fieldList = fieldList & "," & rs2("technical_data_field") & "=null"
											Else
												fieldList = fieldList & "," & rs2("technical_data_field") & "=" & rs3(0)
											End If
										End If
									End If
								End If
							End If
							rs4.Close
						End If
						rs3.Close
						rs2.MoveNext
					Loop
					rs2.Close
					If itemID = 0 Then
						'First add the item to the equipment_items table.
						sqlString = "INSERT INTO equipment_items " & _
								"(equipment_item_name,equipment_item_tag,equipment_type_id,area) " & _
								"VALUES ('" & rs("instr_desc") & "','" & rs("tag") & "',1,'" & rs("plant_area_num") & "')"
						cnInspections.Execute(sqlString)
						sqlString = "SELECT LAST_INSERT_ID()"
						Set rs2 = cnInspections.Execute(sqlString)
						sqlString = ""
						If Not rs2.BOF Then
							rs2.MoveFirst
							If fieldList = "" Then
								fieldList = "equipment_item_id"
							Else
								fieldList = fieldList & ",equipment_item_id"
							End If
							If valueList = "" Then
								valueList = rs2(0)
							Else
								valueList = valueList & "," & rs2(0)
							End If
							'Then insert the data into the psv_technical_data table.
							sqlString = "INSERT INTO psv_technical_data " & _
									"(" & fieldList & ") VALUES (" & valueList & ")"
						End If
					Else
						'Update the existing record in the psv_technical_data table.
						sqlString = "UPDATE psv_technical_data SET " & _
								fieldList & _
								" WHERE equipment_item_id=" & itemID
					End If
					If sqlString <> "" Then
						cnInspections.Execute(sqlString)
					End If
				End If
				rs.MoveNext
			Loop
			Response.Write "Processing Complete<br />"
		End If
		
		Set rs = Nothing
		Set rs2 = Nothing
		Set rs3 = Nothing
		Set rs4 = Nothing
		cnInspections.Close
		Set cnInspections = Nothing
		cnInstruments.Close
		Set cnInstruments = Nothing
		cnInfo.Close
		Set cnInfo = Nothing
	End If
Else
	Response.Write "<h1>You don't have permission to access this page.</h1>"
	Response.Write "<br />"
	Response.Write "<a href='" & request("http_referer") & "'>Return to previous page</a>"
End If

%>
</form>
</body>
</html>
