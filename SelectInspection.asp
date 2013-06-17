<%@ Language=VBScript %>
<% OPTION EXPLICIT %>
<html>
<head>
<script language="javascript">
function editinspection(id,type,cv) {
 if (id=='') {
  alert('You must first select an item to view/edit.');
  } else {
   if (type.toUpperCase()=='PSV') {
    if (cv==true) {
     window.location.href='cv_inspection.asp?inspectionID='+id+'&edit=true';
    } else {
     window.location.href='psv_inspection.asp?inspectionID='+id+'&edit=true';
    }
   } else if (type.toUpperCase()=='PSE') {
     window.location.href='pse_inspection.asp?inspectionID='+id+'&edit=true';
   } else if (type.toUpperCase()=='TANK') {
     window.location.href='tank_inspection.asp?inspectionID='+id+'&edit=true';
   } else {
    alert('Equipment type not specified.');
   }
  }
}
function enterinspection(item,type,cv) {
 if (item=='') {
  alert('Equipment item not specified.');
 } else {
  if (type.toUpperCase()=='PSV') {
   if (cv==true) {
    window.location.href='cv_inspection.asp?itemID='+item+'&edit=true';
   } else {
    window.location.href='psv_inspection.asp?itemID='+item+'&edit=true';
   }
  } else if (type.toUpperCase()=='PSE') {
   window.location.href='pse_inspection.asp?itemID='+item+'&edit=true';
  } else if (type.toUpperCase()=='TANK') {
   window.location.href='tank_inspection.asp?itemID='+item+'&edit=true';
  } else if (type.toUpperCase()=='HEX') {
   window.location.href='hex_inspection.asp?itemID='+item+'&edit=true';
  } else {
   alert('Equipment type not specified.');
  }
 }
}
function manualinspection(item,type,cv) {
 if (item=='') {
  alert('Equipment item not specified.');
 } else {
  if (type.toUpperCase()=='PSV') {
   if (cv==true) {
    window.open('cv_manual_inspection.asp?itemID='+item,'PSVInspection');
   } else {
    window.open('psv_manual_inspection.asp?itemID='+item,'PSVInspection');
   }
  } else if (type.toUpperCase()=='PSE') {
   window.open('pse_manual_inspection.asp?itemID='+item,'PSEInspection');
  } else if (type.toUpperCase()=='TANK') {
   window.open('tank_manual_inspection.asp?itemID='+item,'TankInspection');
  } else if (type.toUpperCase()=='HEX') {
   window.open('hex_manual_inspection.asp?itemID='+item,'HEXInspection');
  } else {
   alert('Equipment type not specified.');
  }
 }
}
function printinspection(id,type,cv) {
 if (id=='') {
  alert('You must first select an item to print.');
  } else {
   if (type.toUpperCase()=='PSV') {
    if (cv==true) {
     window.open('cv_inspection.asp?inspectionID='+id+'&edit=false&print=true','PSVInspection');
    } else {
     window.open('psv_inspection.asp?inspectionID='+id+'&edit=false&print=true','PSVInspection');
    }
   } else if (type.toUpperCase()=='PSE') {
     window.open('pse_inspection.asp?inspectionID='+id+'&edit=false&print=true','PSEInspection');
   } else if (type.toUpperCase()=='TANK') {
     window.open('tank_inspection.asp?inspectionID='+id+'&edit=false&print=true','TankInspection');
   } else if (type.toUpperCase()=='HEX') {
     window.open('hex_inspection.asp?inspectionID='+id+'&edit=false&print=true','HEXInspection');
   } else {
    alert('Equipment type not specified.');
   }
  }
}
function printtechdata(item,type,cv) {
 if (item=='') {
  alert('Equipment item not specified.');
  } else {
   if (type.toUpperCase()=='PSV') {
    if (cv==true) {
     window.open('cv_technicaldata.asp?itemID='+item+'&edit=false&print=true','PSVTechnicalData');
    } else {
     window.open('psv_technicaldata.asp?itemID='+item+'&edit=false&print=true','PSVTechnicalData');
    }
   } else if (type.toUpperCase()=='PSE') {
     window.open('pse_technicaldata.asp?itemID='+item+'&edit=false&print=true','PSETechnicalData');
   } else if (type.toUpperCase()=='TANK') {
     window.open('tank_technicaldata.asp?itemID='+item+'&edit=false&print=true','TankTechnicalData');
   } else if (type.toUpperCase()=='HEX') {
     window.open('hex_technicaldata.asp?itemID='+item+'&edit=false&print=true','HEXTechnicalData');
   } else {
    alert('Equipment type not specified.');
   }
  }
}
function openhelp() {
 window.open("Equipment Inspection Database Users Guide.doc","userguide");
}
</script>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Select Inspection</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="InspectionsFunctions.asp"-->
<link rel=STYLESHEET href='formstyle.css' type='text/css'>
</head>
<body>
<form id="form1" name="form1" action="SelectInspection.asp" method="post">
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
Dim itemType
Dim cv
Dim currentuser
Dim access
Dim access2
Dim itemID
Dim found

'Clear the focus session variable, so the next entry form will start at the
'top field.
Session("focus") = ""

'Store the selection for the action that this form is for so we don't lose the
'querystring value.
Response.Write "<input type='hidden' id='form_action' name='form_action' value='" & Request("form_action") & "' />"

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("inspections", "selectinspection", currentuser)
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
	If Request("itemID") <> "" Then
		'Get the tag of the equipment item for the header.
		sqlString = "SELECT equipment_item_tag,equipment_type_name,conservation_vent " & _
				"FROM equipment_items a INNER JOIN equipment_types b " & _
				"ON a.equipment_type_id=b.equipment_type_id " & _
				"WHERE equipment_item_id=" & Request("itemID")
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			tagname = rs(0)
			itemType = rs(1)
			cv = rs(2)
		End If
		rs.Close

		'Specify hidden fields to store data.
		Response.Write "<input type='hidden' id='item_type' name='item_type' value='" & itemType & "'/>"
		Response.Write "<input type='hidden' id='itemID' name='itemID' value='" & Request("itemID") & "'/>"
		Response.Write "<input type='hidden' id='cv' name='cv' value='" & cv & "'/>"

		'Write the title and header links.
		Response.Write "<table style='width:100%;border:none'>"
		Response.Write "<tr>"
		Response.Write "<td style='text-align:left;vertical-align:top;width:20%'><a href='default.asp'>Home</a></td>"
		Response.Write "<td style='text-align:center;width:60%'><h1 />Inspection for " & tagname & "</td>"
		Response.Write "<td style='text-align:right;vertical-align:top;width:20%'><a href='' onclick='openhelp();return false;' title='Open the User Guide'>Help</a></td>"
		Response.Write "</tr>"
		Response.Write "</table>"
		
'		Response.Write "<div style='text-align:center'>"
'		Response.Write "<h3>Select inspection to view or modify and click the 'View/Edit Inspection' button, click the 'New Inspection' button to create a new inspection for this equipment item using the on-line form, or click the 'Inspection Form' button to print a blank inspection form for this item.</h3>"
'		Response.Write "</div>"
		Response.Write "<br />"
		
		Response.Write "<table style='width:100%;border:none'>"
		Response.Write "<tr>"
		Response.Write "<td style='width:60%'>"

		sqlString = "SELECT inspection_data_id,inspection_date,inspected_by " & _
				"FROM " & LCase(itemType) & "_inspection_data " & _
				"WHERE equipment_item_id=" & Request("itemID") & _
				" ORDER BY inspection_date DESC"
		Set rs = cn.Execute(sqlString)
		found = False
		If Not rs.BOF Then
			found = True
			Response.Write "<table style='width:700px'>"
			Response.Write "<tr>"
			Response.Write "<th style='width:35%'>Inspection Date</th>"
			Response.Write "<th style='width:27%'>Inspected By</th>"
			Response.Write "<th style='width:38%'>&nbsp;</th>"
			Response.Write "</tr>"
			Response.Write "<tr>"
			Response.Write "<td colspan='3'><select id='inspectid' name='inspectid' size='17' style='font-family:courier new' onDblClick='editinspection(document.form1.inspectid.value,document.form1.item_type.value,document.form1.cv.value);'>"
			rs.MoveFirst
			Do While Not rs.EOF
				Dim inspectdate
				Dim inspectby
				If Not IsNull(rs(1)) Then
					inspectdate = "&nbsp;&nbsp;&nbsp;&nbsp;" & Replace(PadRight(FormatDateTime(rs(1),2),13)," ","&nbsp;")
				Else
					inspectdate = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				End If
				If Not IsNull(rs(2)) Then
					inspectby = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & Replace(PadRight(rs(2),42)," ","&nbsp;")
				Else
					inspectby = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & _
							"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & _
							"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				End If
				Response.Write "<option value='" & rs(0) & "'>" & inspectdate & "&nbsp;&nbsp;&nbsp;&nbsp; " & inspectby & "&nbsp;&nbsp; "
				rs.MoveNext
			Loop
			Response.Write "</select></td>"
			Response.Write "</tr>"
			Response.Write "</table>"
			Response.Write "</td>"
		Else
			Response.Write "<h2>No previous inspections found</h2>"
		End If
		rs.Close
		
		Response.Write "<td>"
		Response.Write "<table style='width:100%;border:none'>"
		If found = True Then
			Response.Write "<tr><td>&nbsp;</td></tr>"
			Response.Write "<tr>"
			Response.Write "<td style='text-align:center'><button type='button' style='width:180px' name='editbutton' id='editbutton' title='View/Edit the selected inspection using the on-line form' onclick='editinspection(document.form1.inspectid.value,document.form1.item_type.value,document.form1.cv.value);'>View/Edit Inspection</button></td>"
			Response.Write "</tr>"
		End If
		access2 = UserAccess("inspections", itemType & "_inspection", currentuser)
		If access2 = "write" Or access2 = "delete" Then
			Response.Write "<tr><td>&nbsp;</td></tr>"
			Response.Write "<tr>"
			Response.Write "<td style='text-align:center'><button type='button' style='width:180px' name='enterbutton' id='enterbutton' title='Enter a new inspection for this item using the on-line form' onclick='enterinspection(document.form1.itemID.value,document.form1.item_type.value,document.form1.cv.value);'>Enter New Inspection</button></td>"
			Response.Write "</tr>"
		End If
		Response.Write "<tr><td>&nbsp;</td></tr>"
		Response.Write "<tr>"
		Response.Write "<td style='text-align:center'><button type='button' style='width:180px' name='manualbutton' id='manualbutton' title='Print a blank inspection form for this item' onclick='manualinspection(document.form1.itemID.value,document.form1.item_type.value,document.form1.cv.value);'>Print Inspection Form</button></td>"
		Response.Write "</tr>"
		If found = True Then
			Response.Write "<tr><td>&nbsp;</td></tr>"
			Response.Write "<tr>"
			Response.Write "<td style='text-align:center'><button type='button' style='width:180px' name='printbutton' id='printbutton' title='Print the data for the selected inspection' onclick='printinspection(document.form1.inspectid.value,document.form1.item_type.value,document.form1.cv.value);'>Print Inspection</button></td>"
			Response.Write "</tr>"
		End If
		Response.Write "<tr><td>&nbsp;</td></tr>"
		Response.Write "<tr>"
		Response.Write "<td style='text-align:center'><button type='button' style='width:180px' name='printtechbutton' id='printtechbutton' title='Print the technical data for this item' onclick='printtechdata(document.form1.itemID.value,document.form1.item_type.value,document.form1.cv.value);'>Print Technical Data</button></td>"
		Response.Write "</tr>"
		Response.Write "</table>"

	Else
		Response.Write "<h1>No equipment item specified!</h1>"
	End If

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
