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
 window.open("Equipment Inspection Database Administrators Guide.doc","userguide");
}
function reloadPage() {
 document.form1.action="equipmentitems.asp";
 document.form1.submit()
}
</script>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Administer Equipment Items</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="InspectionsFunctions.asp"-->
<link rel=STYLESHEET href='formstyle.css' type='text/css'>
</head>

<%
'*************
' Revision History
' 
' Keith Brooks - Tuesday, February 15, 2011
'   Creation.
'*************

dim cn
dim rs
Dim rs2
dim recordid
Dim currentuser
Dim access
Dim recid
Dim sqlString
Dim sortkey
Dim sortdir
Dim limitnum

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("inspections","equipmentitems",currentuser)
If access <> "none" Then

	set cn = CreateObject("adodb.connection")
	cn.Open = DBString
	set rs = CreateObject("adodb.recordset")
	set rs2 = CreateObject("adodb.recordset")

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
		sortkey = "equipment_item_id"
	End If
	If request("direction") <> "" Then
		sortdir = request("direction")
	Else
		sortdir = "ASC"
	End If
	If Request("limit") <> "" Then
		limitnum = Request("limit")
	Else
		limitnum = "100"
	End If

	response.write "<table ID='headertable' width='100%'>"
	response.write "<tr>"
	response.write "<td ID='headertd' style='width:20%;text-align:left;vertical-align:top'><a href='adminmenu.asp' title='Open the administration main menu'>Menu</a></td>"
	response.write "<td ID='headertd' style='width:60%;text-align:center;vertical-align:center'><h1/>Edit Equipment Items</td>"
	response.write "<td ID='headertd' style='width:20%;text-align:right;vertical-align:top'><a href='#' onClick='openhelp();return false;' title='Open the Admin Guide'>Help</a></td>"
	response.write "</tr>"

	response.Write "<form action='adminaction.asp' id='form1' method='post' name='form1'>"

	response.write "<tr>"
	response.write "<td id='headertd'>&nbsp;</td>"
	If access = "write" Or access = "delete" Then
		response.write "<td ID='headertd' style='text-align:center'><a href='equipmentitems.asp?record_id=-1&sort=" & sortkey & "&direction=" & sortdir & "&limit=" & limitnum & "' title='Add a new equipment item record'>Add new record</a></td>"
	Else
		Response.Write "<td id='headertd'>&nbsp;</td>"
	End If
	response.write "<td id='headertd'>Records to display:"
	Response.Write "<select name='limit' id='limit' onchange='reloadPage();'>"
	If limitnum = "All" Then
		Response.Write "<option value='All' selected>All"
	Else
		Response.Write "<option value='All'>All"
	End If
	If limitnum = "20" Then
		Response.Write "<option value='20' selected>20"
	Else
		Response.Write "<option value='20'>20"
	End If
	If limitnum = "100" Then
		Response.Write "<option value='100' selected>100"
	Else
		Response.Write "<option value='100'>100"
	End If
	If limitnum = "1000" Then
		Response.Write "<option value='1000' selected>1000"
	Else
		Response.Write "<option value='1000'>1000"
	End If
	Response.Write "</select></td>"
	response.write "</tr>"
	response.write "</table>"

	Response.Write "<br />"

	'Draw the header
	Response.Write "<div style='text-align:center'>"
	response.Write "<table width='100%'>"
	response.Write "<tr>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='equipmentitems.asp?sort=equipment_item_id&direction=DESC&limit=" & limitNum & "'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'&nbsp;>Equipment<br />Item ID&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='equipmentitems.asp?sort=equipment_item_id&direction=ASC&limit=" & limitNum & "'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='equipmentitems.asp?sort=equipment_item_name&direction=DESC&limit=" & limitNum & "'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Equipment<br />Item Name&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='equipmentitems.asp?sort=equipment_item_name&direction=ASC&limit=" & limitNum & "'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='equipmentitems.asp?sort=equipment_item_tag&direction=DESC&limit=" & limitNum & "'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Equipment<br />Item Tag&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='equipmentitems.asp?sort=equipment_item_tag&direction=ASC&limit=" & limitNum & "'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='equipmentitems.asp?sort=equipment_item_description&direction=DESC&limit=" & limitNum & "'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Equipment<br />Item Description&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='equipmentitems.asp?sort=equipment_item_desc&direction=ASC&limit=" & limitNum & "'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='equipmentitems.asp?sort=equipment_type&direction=DESC&limit=" & limitNum & "'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Equipment<br />Type&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='equipmentitems.asp?sort=equipment_type&direction=ASC&limit=" & limitNum & "'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='equipmentitems.asp?sort=conservation_vent&direction=DESC&limit=" & limitNum & "'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Conservation<br />Vent&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='equipmentitems.asp?sort=conservation_vent&direction=ASC&limit=" & limitNum & "'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='equipmentitems.asp?sort=assembly&direction=DESC&limit=" & limitNum & "'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Assembly&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='equipmentitems.asp?sort=assembly&direction=ASC&limit=" & limitNum & "'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='equipmentitems.asp?sort=area&direction=DESC&limit=" & limitNum & "'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Area&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='equipmentitems.asp?sort=area&direction=ASC&limit=" & limitNum & "'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"

	If access = "write" Or access = "delete" Then
 	  response.Write "<th id='mediumth'>&nbsp;</th>"
 	End If
	If access = "delete" Or recordid < 0 Then
 	  response.Write "<th id='mediumth'>&nbsp;</th>"
 	End If
	response.Write "</tr>"

	response.write "<input type='hidden' name='RECORD' value='" & recordid & "'>"
	response.write "<input type='hidden' name='SORT' value='" & sortkey & "'>"
	response.write "<input type='hidden' name='DIRECTION' value='" & sortdir & "'>"

	'Read the form data.
	If limitNum = "All" Then
		sqlString = "SELECT equipment_item_id,area,equipment_item_tag,assembly,conservation_vent," & _
				"equipment_item_name,equipment_item_description,equipment_type_name AS equipment_type " & _
				"FROM equipment_items LEFT JOIN equipment_types " & _
				"ON equipment_items.equipment_type_id=equipment_types.equipment_type_id " & _
				"ORDER BY " & sortkey & " " & sortdir
	Else
		sqlString = "SELECT equipment_item_id,area,equipment_item_tag,assembly,conservation_vent," & _
				"equipment_item_name,equipment_item_description,equipment_type_name AS equipment_type " & _
				"FROM equipment_items LEFT JOIN equipment_types " & _
				"ON equipment_items.equipment_type_id=equipment_types.equipment_type_id " & _
				"ORDER BY " & sortkey & " " & sortdir & " LIMIT " & limitnum
	End If

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

			If session("err") = "equipment_item_name" Then
				If session("equipment_item_name") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='equipment_item_name' name='equipment_item_name' size='50' value='" & session("equipment_item_name") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='equipment_item_name' name='equipment_item_name' size='50' value=''></td>"
				End If
			Else
				If session("equipment_item_name") <> "" Then
					response.write "<td id='mediumtd'><input type='text' class='text' id='equipment_item_name' name='equipment_item_name' size='50' value='" & session("equipment_item_name") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' class='text' id='equipment_item_name' name='equipment_item_name' size='50' value=''></td>"
				End If
			End If

			If session("err") = "equipment_item_tag" Then
				If session("equipment_item_tag") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='equipment_item_tag' name='equipment_item_name' size='15' value='" & session("equipment_item_name") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='equipment_item_tag' name='equipment_item_name' size='15' value=''></td>"
				End If
			Else
				If session("equipment_item_tag") <> "" Then
					response.write "<td id='mediumtd'><input type='text' class='text' id='equipment_item_tag' name='equipment_item_tag' size='15' value='" & session("equipment_item_tag") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' class='text' id='equipment_item_tag' name='equipment_item_tag' size='15' value=''></td>"
				End If
			End If

			If session("err") = "equipment_item_description" Then
				If session("equipment_item_description") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><textarea id='equipment_item_description' name='equipment_item_description' cols='30' rows='2'>" & session("equipment_item_description") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><textarea id='equipment_item_description' name='equipment_item_description' cols='30' rows='2'></textarea></td>"
				End If
			Else
				If session("equipment_item_description") <> "" Then
					response.write "<td id='mediumtd'><textarea id='equipment_item_description' name='equipment_item_description' cols='30' rows='2'>" & session("equipment_item_description") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd'><textarea id='equipment_item_description' name='equipment_item_description' cols='30' rows='2'></textarea></td>"
				End If
			End If

			'Dropdown for equipment type.
			If session("err") = "equipment_type_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select id='equipment_type_id' name='equipment_type_id'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select id='equipment_type_id' name='equipment_type_id'>"
			End If
			sqlString = "SELECT equipment_type_id,equipment_type_name FROM equipment_types ORDER BY equipment_type_name"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("equipment_type_id") <> "" Then
						If CInt(Request("equipment_type_id")) = rs2("equipment_type_id") Then
							response.write "<option value='" & rs2("equipment_type_id") & "' selected>" & rs2("equipment_type_name")
						Else
							response.write "<option value='" & rs2("equipment_type_id") & "'>" & rs2("equipment_type_name")
						End If
					Else
						If Session("equipment_type_id") <> "" Then
							If CInt(session("equipment_type_id")) = rs2("equipment_type_id") Then
								response.write "<option value='" & rs2("equipment_type_id") & "' selected>" & rs2("equipment_type_name")
							Else
								response.write "<option value='" & rs2("equipment_type_id") & "'>" & rs2("equipment_type_name")
							End If
						Else
							response.write "<option value='" & rs2("equipment_type_id") & "'>" & rs2("equipment_type_name")
						End If
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			response.write "</select></td>"

			If session("err") = "conservation_vent" Then
				If Session("conservation_vent") = "" Or session("conservation_vent") = "0" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='checkbox' class='checkbox' id='conservation_vent' name='conservation_vent' value='1' /></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='checkbox' class='checkbox' id='conservation_vent' name='conservation_vent' value='1' checked /></td>"
				End If
			Else
				If session("conservation_vent") = "" Or Session("conservation_vent") = "0" Then
					response.write "<td id='mediumtd'><input type='checkbox' class='checkbox' id='conservation_vent' name='conservation_vent' value='1' /></td>"
				Else
					response.write "<td id='mediumtd'><input type='checkbox' class='checkbox' id='conservation_vent' name='conservation_vent' value='1' checked /></td>"
				End If
			End If

			If session("err") = "assembly" Then
				If session("assembly") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='assembly' name='assembly' size='15' value='" & session("assembly") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='assembly' name='assembly' size='15' value=''></td>"
				End If
			Else
				If session("assembly") <> "" Then
					response.write "<td id='mediumtd'><input type='text' class='text' id='assembly' name='assembly' size='15' value='" & session("assembly") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' class='text' id='assembly' name='assembly' size='15' value=''></td>"
				End If
			End If

			If session("err") = "area" Then
				If session("area") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='area' name='area' size='5' value='" & session("area") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='area' name='area' size='5' value=''></td>"
				End If
			Else
				If session("area") <> "" Then
					response.write "<td id='mediumtd'><input type='text' class='text' id='area' name='area' size='5' value='" & session("area") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' class='text' id='area' name='area' size='5' value=''></td>"
				End If
			End If

			response.write "<td id='mediumtd' style='text-align:center'><input type='submit' value='Submit' id='submit1' name='submit1' style='font-size:8pt;background-color:white' title='Apply changes to this record'></td>"

			Response.Write "<td id='mediumtd' style='text-align:center'><a href='equipmentitems.asp?sort=" & sortkey & "&direction=" & sortdir & "&limit=" & limitnum & "' title='Cancel changes to this record'>Cancel</a></td>"

			Response.Write "</tr>"
		End If
	End If

	Do While Not rs.EOF
		Response.Write "<tr>"
		If CLng(rs("equipment_item_id")) = CLng(recordid) Then
			'Draw the data entry line
			response.write "<td id='mediumtd'>" & rs("equipment_item_id") & "</td>"

			If session("err") = "equipment_item_name" Then
				If session("equipment_item_name") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='equipment_item_name' name='equipment_item_name' size='50' value='" & session("equipment_item_name") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='equipment_item_name' name='equipment_item_name' size='50' value='" & rs("equipment_item_name") & "'></td>"
				End If
			Else
				If session("equipment_item_name") <> "" Then
					response.write "<td id='mediumtd'><input type='text' class='text' id='equipment_item_name' name='equipment_item_name' size='50' value='" & session("equipment_item_name") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' class='text' id='equipment_item_name' name='equipment_item_name' size='50' value='" & rs("equipment_item_name") & "'></td>"
				End If
			End If

			If session("err") = "equipment_item_tag" Then
				If session("equipment_item_tag") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='equipment_item_tag' name='equipment_item_tag' size='15' value='" & session("equipment_item_tag") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='equipment_item_tag' name='equipment_item_tag' size='15' value='" & rs("equipment_item_tag") & "'></td>"
				End If
			Else
				If session("equipment_item_tag") <> "" Then
					response.write "<td id='mediumtd'><input type='text' class='text' id='equipment_item_tag' name='equipment_item_tag' size='15' value='" & session("equipment_item_tag") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' class='text' id='equipment_item_tag' name='equipment_item_tag' size='15' value='" & rs("equipment_item_tag") & "'></td>"
				End If
			End If

			If session("err") = "equipment_item_description" Then
				If session("equipment_item_description") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><textarea id='equipment_item_description' name='equipment_item_description' cols='30' rows='2'>" & session("equipment_item_description") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><textarea id='equipment_item_description' name='equipment_item_description' cols='30' rows='2'>" & rs("equipment_item_description") & "</textarea></td>"
				End If
			Else
				If session("equipment_item_description") <> "" Then
					response.write "<td id='mediumtd'><textarea id='equipment_item_description' name='equipment_item_description' cols='30' rows='2'>" & session("equipment_item_description") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd'><textarea id='equipment_item_description' name='equipment_item_description' cols='30' rows='2'>" & rs("equipment_item_description") & "</textarea></td>"
				End If
			End If

			'Dropdown for equipment type.
			If session("err") = "equipment_type_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select id='equipment_type_id' name='equipment_type_id'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select id='equipment_type_id' name='equipment_type_id'>"
			End If
			sqlString = "SELECT equipment_type_id,equipment_type_name FROM equipment_types ORDER BY equipment_type_name"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("equipment_type_id") <> "" Then
						If CInt(Request("equipment_type_id")) = rs2("equipment_type_id") Then
							response.write "<option value='" & rs2("equipment_type_id") & "' selected>" & rs2("equipment_type_name")
						Else
							response.write "<option value='" & rs2("equipment_type_id") & "'>" & rs2("equipment_type_name")
						End If
					Else
						If rs("equipment_type") = rs2("equipment_type_name") Then
							response.write "<option value='" & rs2("equipment_type_id") & "' selected>" & rs2("equipment_type_name")
						Else
							response.write "<option value='" & rs2("equipment_type_id") & "'>" & rs2("equipment_type_name")
						End If
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			response.write "</select></td>"

			If session("err") = "conservation_vent" Then
				If rs("conservation_vent") = 0 Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='checkbox' class='checkbox' id='conservation_vent' name='conservation_vent' value='1' /></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='checkbox' class='checkbox' id='conservation_vent' name='conservation_vent' value='1' checked /></td>"
				End If
			Else
				If rs("conservation_vent") = 0 Then
					response.write "<td id='mediumtd'><input type='checkbox' class='checkbox' id='conservation_vent' name='conservation_vent' value='1' /></td>"
				Else
					response.write "<td id='mediumtd'><input type='checkbox' class='checkbox' id='conservation_vent' name='conservation_vent' value='1' checked /></td>"
				End If
			End If

			If session("err") = "assembly" Then
				If session("assembly") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='assembly' name='assembly' size='15' value='" & session("assembly") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='assembly' name='assembly' size='15' value='" & rs("assembly") & "'></td>"
				End If
			Else
				If session("assembly") <> "" Then
					response.write "<td id='mediumtd'><input type='text' class='text' id='assembly' name='assembly' size='15' value='" & session("assembly") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' class='text' id='assembly' name='assembly' size='15' value='" & rs("assembly") & "'></td>"
				End If
			End If

			If session("err") = "area" Then
				If session("area") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='area' name='area' size='5' value='" & session("area") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' class='text' id='area' name='area' size='5' value='" & rs("area") & "'></td>"
				End If
			Else
				If session("area") <> "" Then
					response.write "<td id='mediumtd'><input type='text' class='text' id='area' name='area' size='5' value='" & session("area") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' class='text' id='area' name='area' size='5' value='" & rs("area") & "'></td>"
				End If
			End If

			If access = "write" Or access = "delete" Then
				response.write "<td id='mediumtd' style='text-align:center'><input type='submit' value='Submit' id='submit1' name='submit1' style='font-size:8pt;background-color:white' title='Apply changes to this record'></td>"
			End If

			If access = "delete" Then
				recid = rs("equipment_item_id")
				response.write "<td id='mediumtd'><a href='javascript:doDelete(" & recid & ");' title='Delete this record'>Delete</a></td>"
			End If
		Else
			'Draw the history records
			response.write "<tr>"
			response.write "<td id='mediumtd'>" & rs("equipment_item_id") & "</td>"
			If rs("equipment_item_name") <> "" Then
				response.write "<td id='mediumtd'>" & rs("equipment_item_name") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If rs("equipment_item_tag") <> "" Then
				response.write "<td id='mediumtd'>" & rs("equipment_item_tag") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If rs("equipment_item_description") <> "" And rs("equipment_item_description") <> " " Then
				response.write "<td id='mediumtd'>" & rs("equipment_item_description") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If rs("equipment_type") <> "" Then
				response.write "<td id='mediumtd'>" & rs("equipment_type") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If rs("conservation_vent") = 0 Then
				Response.write "<td id='mediumtd'><input type='checkbox' class='checkbox' disabled /></td>"
			Else
				Response.write "<td id='mediumtd'><input type='checkbox' class='checkbox' checked disabled /></td>"
			End If
			If rs("assembly") <> "" Then
				response.write "<td id='mediumtd'>" & rs("assembly") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If rs("area") <> "" Then
				response.write "<td id='mediumtd'>" & rs("area") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If access = "write" Or access = "delete" Then
				response.write "<td id='mediumtd'><a href='equipmentitems.asp?record_id=" & rs("equipment_item_id") & "&sort=" & sortkey & "&direction=" & sortdir & "&limit=" & limitnum & "' title='Edit this record'>Edit</a></td>"
			End If
			If access = "delete" Then
				recid = rs("equipment_item_id")
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
