<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<html>
<head>
<script language="javascript">
function doSubmit() {
 document.getElementById('PleaseWait').style.display = 'block';
 document.form1.flowflag.value='false';
 document.form1.submit();
}
function doFind() {
 document.getElementById('PleaseWait').style.display = 'block';
 document.form1.submit();
}
function editdata(id) {
 if (id=='') {
  alert('You must select an item to edit.');
  } else {
   document.form1.action='technicaldataaction.asp?itemID='+id;
   document.form1.submit();
  }
}
function inspect(id) {
 if (id=='') {
  alert('You must select an item for the inspection.');
  } else {
  window.location.href='selectinspection.asp?itemID='+id;
  }
}
function openhelp() {
 window.open("Equipment Inspection Database Users Guide.doc","userguide");
}
</script>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Select Equipment Item</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="InspectionsFunctions.asp"-->
<link rel=STYLESHEET href='formstyle.css' type='text/css'>
</head>
<body>
<table style="width:100%;border:none">
	<tr>
		<td style="text-align:left;vertical-align:top;width:20%"><a href="default.asp">Home</a></td>
		<td style="text-align:center;width:60%"><h1 />Select Equipment Item</td>
		<td style="text-align:right;vertical-align:top;width:20%"><a href="" onclick="openhelp();return false;" title="Open the User Guide">Help</a></td>
	</tr>
</table>
<form id="form1" name="form1" action="SelectItem.asp" method="post">
<%
'*************
' Revision History
' 
' Keith Brooks - Monday, January 17, 2011
'   Creation.
'*************

Dim sqlString
Dim cn
Dim rs
Dim criteria
Dim tagname
Dim tagdesc
Dim currentuser
Dim access
Dim itemID

'Store the selection for the action that this form is for so we don't lose the
'querystring value.
Response.Write "<input type='hidden' id='form_action' name='form_action' value='" & Request("form_action") & "' />"

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("inspections", "selectitem", currentuser)
If access <> "none" Then

	'Define the ado connection and recordset objects.
	set cn = CreateObject("adodb.connection")
	cn.Open = DBString
	set rs = CreateObject("adodb.recordset")

	'Draw "Please Wait..." message that will be displayed when this page is
	'reloading, saving data, or moving to another page.
	%>
		<div class="helptext" id="PleaseWait" style="display: none; text-align:center; color:White; vertical-align:top;border-style:none;position:absolute;top:0px;left:0px">
			<table id="MyTable" bgcolor="blue">
				<tr><td style="width: 95px"><b><font color="white">Please Wait...</font></b></td></tr>
			</table>
		</div>
	<%
	'Draw the criteria selection lists.
	Response.Write "<div style='text-align:center'>"
	Response.Write "<table style='width:50%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<th style='width:50%'>Plant Area</th>"
	Response.Write "<th style='width:50%'>Equipment Type</th>"
	Response.Write "</tr>"
	Response.Write "<tr>"
	'Load the plant area dropdown list.
	Response.Write "<td><select id='plantarea' name='plantarea'>"
	sqlString = "SELECT DISTINCT area FROM equipment_items ORDER BY area"
	Set rs = cn.Execute(sqlString)
	If Not rs.BOF Then
		rs.MoveFirst
		Response.Write "<option value=''> "
		Do While Not rs.EOF
			If Request("plantarea") <> "" Then
				If rs(0) = Request("plantarea") Then
					Response.Write "<option value='" & rs(0) & "' selected>" & rs(0)
				Else
					Response.Write "<option value='" & rs(0) & "'>" & rs(0)
				End If
			Else
				Response.Write "<option value='" & rs(0) & "'>" & rs(0)
			End If
			rs.MoveNext
		Loop
	End If
	rs.Close
	Response.Write "</select></td>"
	'Load the equipment type dropdown list.
	Response.Write "<td><select id='equiptype' name='equiptype'>"
	sqlString = "SELECT equipment_type_id,equipment_type_name FROM equipment_types ORDER BY equipment_type_name"
	Set rs = cn.Execute(sqlString)
	If Not rs.BOF Then
		rs.MoveFirst
		Response.Write "<option value=''> "
		Do While Not rs.EOF
			If Request("equiptype") <> "" Then
				If CLng(rs(0)) = CLng(Request("equiptype")) Then
					Response.Write "<option value='" & rs(0) & "' selected>" & rs(1)
				Else
					Response.Write "<option value='" & rs(0) & "'>" & rs(1)
				End If
			Else
				Response.Write "<option value='" & rs(0) & "'>" & rs(1)
			End If
			rs.MoveNext
		Loop
	End If
	rs.Close
	Response.Write "</select></td>"
	Response.Write "</tr>"

	Response.Write "</table>"
	Response.Write "</div>"

	Response.Write "<br />"
	Response.Write "<div style='text-align:center'>"
	Response.Write "<input type='button' id='submit1' name='submit1' value='Find' onclick='doFind();' />"
	
	'If any of the criteria have been selected, display the list box with the results.
	criteria = ""
	If Request("plantarea") <> "" Then
		criteria = "area=" & Request("plantarea")
	End If
	If Request("equiptype") <> "" Then
		If criteria = "" Then
			criteria = "equipment_type_id=" & Request("equiptype")
		Else
			criteria = criteria & " AND equipment_type_id=" & Request("equiptype")
		End If
	End If

	If Request("flowflag") = "true" And criteria <> "" Then
		sqlString = "SELECT equipment_item_id,equipment_item_tag," & _
				"equipment_item_name " & _
				"FROM equipment_items " & _
				"WHERE " & criteria & " ORDER BY equipment_item_tag"
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			Response.Write "<div style='text-align:center'><table style='width:700px'>"
			Response.Write "<tr>"
			Response.Write "<th style='width:30%'>Item Tag</th>"
			Response.Write "<th style='width:10%'>&nbsp;</th>"
			Response.Write "<th style='width:60%;text-align:left'>Item Description</th>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			If Request("form_action") = "edittechnicaldata" Then
				Response.Write "<td colspan='3'><select id='item_id' name='item_id' size='17' style='font-family:courier new' onDblClick='editdata(document.form1.item_id.value);'>"
			Else
				Response.Write "<td colspan='3'><select id='item_id' name='item_id' size='17' style='font-family:courier new' onDblClick='inspect(document.form1.item_id.value);'>"
			End If
			Do While Not rs.EOF
				If Not IsNull(rs(1)) Then
					tagname = Replace(PadRight(rs(1),17)," ","&nbsp;")
				Else
					tagname = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				End If
				If Not IsNull(rs(2)) Then
					tagdesc = Replace(PadRight(rs(2),52)," ","&nbsp;")
				Else
					tagdesc = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & _
							"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & _
							"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				End If
				Response.Write "<option value='" & rs(0) & "'>" & tagname & "&nbsp;&nbsp;&nbsp;&nbsp; " & tagdesc & "&nbsp;&nbsp; "
				rs.MoveNext
			Loop
			rs.Close
			Response.Write "</select></td>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td colspan='3'>"
			Response.Write "<table style='width:100%'>"
			Response.Write "<tr>"
			If Request("form_action") = "edittechnicaldata" And (access = "write" Or access = "delete") Then
				Response.Write "<td style='text-align:center'><button type='button' name='editbutton' id='editbutton' onclick='editdata(document.form1.item_id.value);'>Edit Technical Data</button></td>"
			Else
				Response.Write "<td style='text-align:center'><button type='button' name='inspectbutton' id='inspectbutton' onclick='inspect(document.form1.item_id.value);'>Continue</button></td>"
			End If
			Response.Write "</tr>"
			Response.Write "</table>"
			Response.Write "</td>"
			Response.Write "</tr>"
			Response.Write "</table></div>"
		Else
			Response.Write "<div style='text-align:center'>"
			Response.Write "<h2>No Data Found</h2>"
			Response.Write "</div>"
		End If
	End If

	'This flag is used to prevent premature retrieval of the item data when one or
	'more of the criteria objects performs a submit.
	Response.Write "<input type='hidden' name='flowflag' id='flowflag' value='true' />"

	Set rs = Nothing
	cn.Close
	Set cn = Nothing
Else
	Response.Write "<h1>You don't have permission to access this page.</h1>"
	Response.Write "<br />"
	Response.Write "<a href='" & request("http_referer") & "'>Return to previous page</a>"
End If
%>
</form>
</body>
</html>
