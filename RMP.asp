<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<!--#include file="..\Functions\HitCounter.asp"-->
<html>
<head>
<script language="javascript">
function openhelp() {
 window.open("Equipment Inspection Database Users Guide.doc","userguide");
}
</script>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>RMP</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="InspectionsFunctions.asp"-->
<link rel=STYLESHEET href='formstyle.css' type='text/css'>
</head>
<body>
<%
'*************
' Revision History
' 
' Keith Brooks - Friday, May 20, 2011
'   Creation
'*************

Dim HitCounts
Dim currentuser
Dim access
Dim cn
Dim cnInstr
Dim rs
Dim rsInstr
Dim sqlString
Dim areas()
Dim area
Dim eda_tags()
Dim eda_types()
Dim dorlastan_tags()
Dim dorlastan_types()
Dim roica_tags()
Dim roica_types()
Dim tags()
Dim types()
Dim tag
Dim tagtype
Dim counter

'Initialize arrays.
ReDim areas(2)
areas(0) = "EDA Storage"
areas(1) = "Dorlastan Amines"
areas(2) = "Roica Amines"

ReDim eda_tags(2)
eda_tags(0) = "0242-070"
eda_tags(1) = "0242-07004A"
eda_tags(2) = "0242-07004B"

ReDim eda_types(2)
eda_types(0) = "Tank"
eda_types(1) = "PSV"
eda_types(2) = "PSV"

ReDim dorlastan_tags(11)
dorlastan_tags(0) = "0230-094"
dorlastan_tags(1) = "0230-09402"
dorlastan_tags(2) = "0230-09403"
dorlastan_tags(3) = "0230-070"
dorlastan_tags(4) = "0230-07002"
dorlastan_tags(5) = "0230-072"
dorlastan_tags(6) = "0230-07201"
dorlastan_tags(7) = "0230-07205"
dorlastan_tags(8) = "0230-073"
dorlastan_tags(9) = "0230-07301"
dorlastan_tags(10) = "0230-075"
dorlastan_tags(11) = "0230-07501"

ReDim dorlastan_types(11)
dorlastan_types(0) = "Tank"
dorlastan_types(1) = "PSV"
dorlastan_types(2) = "PSV"
dorlastan_types(3) = "Tank"
dorlastan_types(4) = "PSV"
dorlastan_types(5) = "Tank"
dorlastan_types(6) = "PSV"
dorlastan_types(7) = "PSV"
dorlastan_types(8) = "Tank"
dorlastan_types(9) = "PSV"
dorlastan_types(10) = "Tank"
dorlastan_types(11) = "PSV"

ReDim roica_tags(5)
roica_tags(0) = "0230-320"
roica_tags(1) = "0230-32001"
roica_tags(2) = "0230-32002"
roica_tags(3) = "0230-321"
roica_tags(4) = "0230-32101"
roica_tags(5) = "0230-32102"

ReDim roica_types(5)
roica_types(0) = "Tank"
roica_types(1) = "PSV"
roica_types(2) = "PSV"
roica_types(3) = "Tank"
roica_types(4) = "PSV"
roica_types(5) = "PSV"

'Set/get hit counts.
HitCounts = HitCounter("rmp")

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("inspections", "default", currentuser)
If access <> "none" Then

	'Create the data objects.
	set cn = CreateObject("adodb.connection")
	cn.Open = DBString
	set rs = CreateObject("adodb.recordset")
	set cnInstr = CreateObject("adodb.connection")
	cnInstr.Open = InstrDBString
	set rsInstr = CreateObject("adodb.recordset")
%>
	<div style="text-align:center">
		<table width="100%">
			<tr>
				<td style="width:25%">&nbsp;</td>
				<td style="text-align:center;width:50%"><h1>RMP</h1></td>
				<td style="text-align:right;vertical-align:top;width:25%"><a href="" onclick="openhelp();return false;" title="Open the User Guide">Help</a></td>
			</tr>
		</table>
	</div>
	<table id="headertable" style="width:100%">
		<tr>
			<th id="headerth" style="width:15%">&nbsp;</th>
			<th id="headerth" style="width:10%">&nbsp;</th>
			<th id="headerth" style="width:10%">&nbsp;</th>
			<th id="headerth" style="width:15%">Date Last<br />Inspection</th>
			<th id="headerth" style="width:10%">Freq.<br />Years</th>
			<th id="headerth" style="width:15%">Date Next<br />Inspection</th>
			<th id="headerth" style="width:15%">P&ID</th>
			<th id="headerth" style="width:10%">Standard</td>
		</tr>
<%
	For Each area In areas
%>
		<tr>
			<td id="formtd" colspan="8" style="font-weight:bold;text-decoration:underline"><%=area%></td>
		</tr>
<%
		'Put the appropriate tags and types into the temporary arrays.
		If area = "EDA Storage" Then
			ReDim tags(UBound(eda_tags))
			ReDim types(UBound(eda_tags))
			For counter = 0 To UBound(eda_tags)
				tags(counter) = eda_tags(counter)
				types(counter) = eda_types(counter)
			Next
		ElseIf area = "Dorlastan Amines" Then
			ReDim tags(UBound(dorlastan_tags))
			ReDim types(UBound(dorlastan_tags))
			For counter = 0 To UBound(dorlastan_tags)
				tags(counter) = dorlastan_tags(counter)
				types(counter) = dorlastan_types(counter)
			Next
		ElseIf area = "Roica Amines" Then
			ReDim tags(UBound(roica_tags))
			ReDim types(UBound(roica_tags))
			For counter = 0 To UBound(roica_tags)
				tags(counter) = roica_tags(counter)
				types(counter) = roica_types(counter)
			Next
		End If
		counter = 0
		For Each tag In tags
			'Get the data from the database for this tag.
			Dim temp
			If types(counter) = "PSV" Then
				temp = "PSV-" & tag
			Else
				temp = tag
			End If
			sqlString = "SELECT i.equipment_item_id,t.inspection_frequency," & _
				"t.inspection_frequency_units,t.next_inspection_date," & _
				"t.drawing_number,t.inspection_standard,i.conservation_vent " & _
				"FROM equipment_items i INNER JOIN " & _
				LCase(types(counter)) & "_technical_data t " & _
				"ON i.equipment_item_id=t.equipment_item_id " & _
				"WHERE i.equipment_item_tag='" & temp & "'"
			Set rs = cn.Execute(sqlString)
			Dim id
			Dim last_inspection
			Dim last_inspection_id
			Dim frequency
			Dim next_inspection
			Dim drawing_number
			Dim inspection_standard
			Dim dwg_id
			Dim equipType
			id = 0
			last_inspection = ""
			last_inspection_id = 0
			frequency = ""
			next_inspection = ""
			drawing_number = ""
			inspection_standard = ""
			dwg_id = 0
			equipType = ""
			If Not rs.BOF Then
				rs.MoveFirst
				If Not IsNull(rs("equipment_item_id")) Then
					id = rs("equipment_item_id")
				End If
				If Not IsNull(rs("inspection_frequency")) Then
					If rs("inspection_frequency_units") = "years" Then
						frequency = rs("inspection_frequency")
					ElseIf rs("inspection_frequency_units") = "months" Then
						frequency = rs("inspection_frequency") / 12
					End If
				End If
				If Not IsNull(rs("next_inspection_date")) Then
					next_inspection = rs("next_inspection_date")
				End If
				If Not IsNull(rs("drawing_number")) Then
					drawing_number = rs("drawing_number")
					'Get the drawing id for the drawing.
					sqlString = "SELECT dwg_id FROM drawings " & _
							"WHERE dwg_name='" & drawing_number & "' " & _
							"AND dwg_type_id=1"
					Set rsInstr = cnInstr.Execute(sqlString)
					If Not rsInstr.BOF Then
						rsInstr.MoveFirst
						If Not IsNull(rsInstr(0)) Then
							dwg_id = rsInstr(0)
						End If
					End If
					rsInstr.Close
				End If
				If Not IsNull(rs("inspection_standard")) Then
					inspection_standard = rs("inspection_standard")
				End If
				If rs("conservation_vent") = True Then
					equipType = "cv"
				Else
					equipType = LCase(types(counter))
				End If
			End If
			rs.Close
			'Get the last inspection date.
			sqlString = "SELECT inspection_data_id,inspection_date " & _
					"FROM " & LCase(types(counter)) & "_inspection_data " & _
					"WHERE inspection_date=" & _
					"(SELECT MAX(inspection_date) " & _
					"FROM " & LCase(types(counter)) & "_inspection_data " & _
					"WHERE equipment_item_id=" & id & ")"
			Set rs = cn.Execute(sqlString)
			If Not rs.BOF Then
				rs.MoveFirst
				last_inspection_id = rs(0)
				If Not IsNull(rs(1)) Then
					last_inspection = rs(1)
				End If
			End If
			rs.Close
						
%>
		<tr>
			<td id="formtd"><%=tag%></td>
			<td id="formtd"><%=types(counter)%></td>
			<td id="formtd" style="text-align:center"><a href="selectinspection.asp?itemID=<%=id%>">Archive</a></td>
			<td id="formtd" style="text-align:center"><a href="<%=equipType%>_inspection.asp?inspectionID=<%=last_inspection_id%>"><%=last_inspection%></a></td>
			<td id="formtd" style="text-align:center"><%=frequency%></td>
			<td id="formtd" style="text-align:center"><%=next_inspection%></td>
			<td id="formtd" style="text-align:center"><a href="http://mogsb8/drawings/drawing.asp?dwg_id=<%=dwg_id%>"><%=drawing_number%></a></td>
			<td id="formtd" style="text-align:center"><%=inspection_standard%></td>
		</tr>
<%
			counter = counter + 1
		Next
	Next
%>
	</table>
<%
	Set rs = Nothing
	Set rsInstr = Nothing
	cn.Close
	Set cn = Nothing
	cnInstr.Close
	Set cnInstr = Nothing
Else
	response.write "<h1>You don't have permission to access this page.</h1>"
	response.write "<br />"
	response.write "<a href='" & request("http_referer") & "'>Return to previous page</a>"
End If
%>
</body>
</html>
